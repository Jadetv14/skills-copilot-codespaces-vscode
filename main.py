import sys
import os
import sqlite3
import logging
import shutil
from datetime import datetime
from typing import Optional, List, Tuple, Dict, Any
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QDateTimeEdit, QComboBox, QTabWidget, QStatusBar, QMessageBox, QFileDialog,
    QGroupBox, QSplitter, QCalendarWidget, QSpinBox, QCheckBox, QDateEdit,
    QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QTimer, QDateTime, QDate, QTime, QUrl
from PyQt5.QtGui import QColor, QPalette, QIcon, QPixmap, QRegExpValidator, QRegExp
import obsws_python as obs
from fpdf import FPDF

# Configurações globais
APP_NAME = "OBS Control Pro"
VERSION = "4.1"  # Atualizado para versão com melhorias
DEVELOPER = "Felipe Iglesias"
DB_FILE = "obs_control.db"
LOG_FILE = "obs_control.log"

# Configuração de logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='a'
)
logger = logging.getLogger(__name__)

class OBSController:
    """Gerencia a conexão e comunicação com o OBS Studio via WebSocket."""
    
    def __init__(self):
        self.client = None
        self.connected = False
        self.last_error = ""
        self._observers = []

    def add_observer(self, callback):
        """Adiciona uma função para ser chamada quando a cena mudar."""
        self._observers.append(callback)

    def _notify_observers(self, scene_name):
        """Notifica todos os observadores sobre mudança de cena."""
        for observer in self._observers:
            observer(scene_name)

    def connect(self, host: str, port: str, password: str) -> bool:
        """Estabelece conexão com o OBS Studio."""
        try:
            if not all([host, port, password]):
                self.last_error = "Preencha todos os campos"
                return False

            if not port.isdigit() or not (1 <= int(port) <= 65535):
                self.last_error = "Porta inválida (1-65535)"
                return False

            self.client = obs.ReqClient(
                host=host,
                port=int(port),
                password=password,
                timeout=3
            )
            
            _ = self.client.get_version()
            self.connected = True
            logger.info(f"Conectado ao OBS em {host}:{port}")
            return True

        except Exception as e:
            self.last_error = f"Erro: {str(e)}"
            logger.error(f"Falha na conexão: {str(e)}")
            return False

    def disconnect(self):
        """Desconecta do OBS Studio."""
        if self.connected and self.client:
            try:
                self.client.disconnect()
                logger.info("Desconectado do OBS")
            except Exception as e:
                logger.error(f"Erro ao desconectar: {str(e)}")
            finally:
                self.connected = False

    def get_scenes(self) -> List[str]:
        """Retorna lista de todas as cenas disponíveis."""
        if not self.connected:
            return []

        try:
            scenes = self.client.get_scene_list()
            return [s['sceneName'] for s in scenes.scenes]
        except Exception as e:
            logger.error(f"Erro ao listar cenas: {str(e)}")
            return []

    def get_current_scene(self) -> Optional[str]:
        """Retorna o nome da cena atualmente ativa."""
        if not self.connected:
            return None
            
        try:
            response = self.client.get_current_program_scene()
            return response.current_program_scene_name
        except Exception as e:
            logger.error(f"Erro ao obter cena atual: {str(e)}")
            return None

    def set_scene(self, scene_name: str) -> bool:
        """Muda para a cena especificada."""
        if not self.connected:
            return False

        try:
            self.client.set_current_program_scene(scene_name)
            logger.info(f"Cena alterada: {scene_name}")
            self._notify_observers(scene_name)
            return True
        except Exception as e:
            logger.error(f"Erro ao mudar de cena: {str(e)}")
            return False
            
class DatabaseManager:
    """Gerencia todas as operações com o banco de dados SQLite."""
    
    def __init__(self, db_file: str = DB_FILE):
        self.db_file = db_file
        self.connection = None

    def connect(self) -> bool:
        """Estabelece conexão com o banco de dados e cria tabelas se necessário."""
        try:
            self.connection = sqlite3.connect(self.db_file)
            self._create_tables()
            self._create_indexes()
            self.verify_database_structure()
            logger.info("Banco de dados conectado e pronto")
            return True
        except sqlite3.Error as e:
            logger.error(f"Erro ao conectar ao banco: {str(e)}")
            return False

    def verify_database_structure(self):
        """Verifica e atualiza a estrutura do banco de dados."""
        with self.connection:
            cursor = self.connection.cursor()
            
            # Verifica e adiciona colunas ausentes
            tables = {
                'scene_history': ['status'],
                'clientes': ['agencia_id'],
                'scene_schedule': ['midia_id']
            }
            
            for table, columns in tables.items():
                cursor.execute(f"PRAGMA table_info({table})")
                existing_columns = [col[1] for col in cursor.fetchall()]
                
                for col in columns:
                    if col not in existing_columns:
                        try:
                            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col} INTEGER DEFAULT NULL")
                            logger.info(f"Coluna '{col}' adicionada à tabela {table}")
                        except sqlite3.Error as e:
                            logger.error(f"Erro ao adicionar coluna {col} em {table}: {str(e)}")

    def _create_tables(self):
        """Cria as tabelas necessárias se não existirem."""
        with self.connection:
            cursor = self.connection.cursor()
            
            # Tabela de agendamentos
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS scene_schedule (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    scene_name TEXT NOT NULL,
                    schedule_time TEXT NOT NULL,
                    repeat_days INTEGER DEFAULT 0,
                    notes TEXT,
                    status TEXT DEFAULT 'pending',
                    midia_id INTEGER DEFAULT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(midia_id) REFERENCES midias(id)
                )
            """)

            # Tabela de histórico
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS scene_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    scene_name TEXT NOT NULL,
                    start_time TEXT NOT NULL,
                    end_time TEXT,
                    duration INTEGER,
                    source_file TEXT,
                    status TEXT DEFAULT 'executed',
                    midia_id INTEGER DEFAULT NULL,
                    FOREIGN KEY(midia_id) REFERENCES midias(id)
                )
            """)

            # Tabelas de cadastro
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS emissoras (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    razao_social TEXT NOT NULL,
                    cnpj TEXT UNIQUE NOT NULL,
                    endereco TEXT,
                    contato TEXT,
                    email TEXT,
                    logo_path TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS clientes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    razao_social TEXT NOT NULL,
                    cnpj TEXT UNIQUE NOT NULL,
                    ie TEXT,
                    contato TEXT,
                    agencia_id INTEGER DEFAULT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(agencia_id) REFERENCES agencias(id)
                )
            """)

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS agencias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    razao_social TEXT NOT NULL,
                    cnpj TEXT UNIQUE NOT NULL,
                    ie TEXT,
                    contato TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Nova tabela de mídias
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS midias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    caminho TEXT NOT NULL,
                    cliente_id INTEGER NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(cliente_id) REFERENCES clientes(id)
                )
            """)

    def _create_indexes(self):
        """Cria índices para melhorar performance das consultas."""
        with self.connection:
            cursor = self.connection.cursor()
            indexes = [
                "CREATE INDEX IF NOT EXISTS idx_schedule_time ON scene_schedule(schedule_time)",
                "CREATE INDEX IF NOT EXISTS idx_schedule_status ON scene_schedule(status)",
                "CREATE INDEX IF NOT EXISTS idx_midia_cliente ON midias(cliente_id)",
                "CREATE INDEX IF NOT EXISTS idx_cliente_agencia ON clientes(agencia_id)"
            ]
            
            for index in indexes:
                cursor.execute(index)

    def get_scheduled_scenes(self, datetime_str: str) -> List[Tuple]:
        """Recupera cenas agendadas para o horário especificado."""
        try:
            with self.connection:
                cursor = self.connection.cursor()
                cursor.execute("""
                    SELECT id, scene_name, repeat_days
                    FROM scene_schedule
                    WHERE status = 'pending'
                    AND schedule_time <= ?
                """, (datetime_str,))
                return cursor.fetchall()
        except sqlite3.Error as e:
            logger.error(f"Erro ao buscar cenas agendadas: {str(e)}")
            return []

    # ... (outros métodos do DatabaseManager permanecem iguais)
    
class MainWindow(QMainWindow):
    """Janela principal da aplicação com todas as interfaces."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{VERSION} - {DEVELOPER}")
        self.setGeometry(100, 100, 1200, 800)
        
        self.obs = OBSController()
        self.db = DatabaseManager()
        self.current_scene = None
        self.scene_start_time = None
        self.emissora_logo_path = ""
        self.midia_file_path = ""
        
        # Configura atalhos
        self.setup_shortcuts()
        
        # Primeiro conecta ao banco antes de inicializar a UI
        if not self._initialize_database():
            return
            
        self._init_ui()
        self._setup_timers()
        self.setup_backup()

    def setup_shortcuts(self):
        """Configura atalhos de teclado."""
        self.shortcuts = {
            'connect': QShortcut("Ctrl+C", self),
            'add_schedule': QShortcut("Ctrl+A", self),
            'generate_report': QShortcut("Ctrl+G", self)
        }
        
    def setup_backup(self):
        """Configura sistema de backup automático."""
        self.backup_timer = QTimer()
        self.backup_timer.timeout.connect(self._do_backup)
        self.backup_timer.start(3600000)  # 1 hora

    def _do_backup(self):
        """Executa backup do banco de dados."""
        try:
            if not os.path.exists("backups"):
                os.makedirs("backups")
                
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = f"backups/obs_backup_{timestamp}.db"
            shutil.copy2(DB_FILE, backup_file)
            
            # Mantém apenas os últimos 5 backups
            backups = sorted([f for f in os.listdir("backups") if f.startswith("obs_backup")])
            for old_backup in backups[:-5]:
                os.remove(os.path.join("backups", old_backup))
                
            logger.info(f"Backup criado: {backup_file}")
        except Exception as e:
            logger.error(f"Erro ao criar backup: {str(e)}")

    def _initialize_database(self):
        """Inicializa o banco de dados com tratamento de erros."""
        try:
            if not self.db.connect():
                self._show_error("Falha crítica: Não foi possível conectar ao banco de dados")
                return False
            return True
        except Exception as e:
            self._show_error(f"Erro ao conectar ao banco: {str(e)}")
            return False

    def _init_ui(self):
        """Configura todos os elementos da interface gráfica."""
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Arquivo")
        about_menu = menubar.addMenu("Sobre")
        
        exit_action = file_menu.addAction("Sair")
        exit_action.triggered.connect(self.close)
        
        about_action = about_menu.addAction("Sobre o Sistema")
        about_action.triggered.connect(self._show_about)
        
        self.tabs = QTabWidget()
        self._setup_monitor_tab()
        self._setup_schedule_tab()
        self._setup_reports_tab()
        self._setup_cadastro_tab()
        self.setCentralWidget(self.tabs)
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(f"Pronto para conectar | {DEVELOPER} | v{VERSION}")

    def _setup_timers(self):
        """Configura os timers para atualização automática."""
        self.scene_timer = QTimer()
        self.scene_timer.timeout.connect(self._update_scenes)
        self.scene_timer.start(2000)  # Atualiza a cada 2 segundos

        self.schedule_timer = QTimer()
        self.schedule_timer.timeout.connect(self._check_schedules)
        self.schedule_timer.start(1000)  # Verifica agendamentos a cada 1 segundo

    def _setup_monitor_tab(self):
        """Configura a aba de monitoramento."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Grupo de conexão
        connection_group = QGroupBox("Conexão com OBS")
        connection_layout = QFormLayout(connection_group)
        
        self.host_input = QLineEdit("localhost")
        self.port_input = QLineEdit("4455")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        
        self.connect_btn = QPushButton("Conectar (Ctrl+C)")
        self.connect_btn.clicked.connect(self._connect_to_obs)
        self.shortcuts['connect'].activated.connect(self._connect_to_obs)
        
        connection_layout.addRow("Host:", self.host_input)
        connection_layout.addRow("Porta:", self.port_input)
        connection_layout.addRow("Senha:", self.password_input)
        connection_layout.addRow(self.connect_btn)
        
        # Grupo de cena atual
        self.scene_group = QGroupBox("Cena Atual")
        scene_layout = QVBoxLayout(self.scene_group)
        
        self.current_scene_label = QLabel("Nenhuma cena detectada")
        self.current_scene_label.setAlignment(Qt.AlignCenter)
        self.current_scene_label.setStyleSheet("""
            font-size: 16px; 
            font-weight: bold;
            color: #ffffff;
        """)
        
        # Tabela de cenas
        self.scenes_table = QTableWidget()
        self.scenes_table.setColumnCount(2)
        self.scenes_table.setHorizontalHeaderLabels(["Cena", "Status"])
        self.scenes_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.scenes_table.setSelectionMode(QTableWidget.NoSelection)
        
        scene_layout.addWidget(self.current_scene_label)
        scene_layout.addWidget(self.scenes_table)
        
        layout.addWidget(connection_group)
        layout.addWidget(self.scene_group)
        self.tabs.addTab(tab, "Monitoramento")
        
    def _setup_schedule_tab(self):
        """Configura a aba de agendamento com calendário e agendamento de cenas."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Divisão principal
        splitter = QSplitter(Qt.Horizontal)
        
        # Painel esquerdo - Calendário e formulário
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        # Grupo do calendário
        calendar_group = QGroupBox("Selecione a Data e Hora")
        calendar_layout = QVBoxLayout(calendar_group)
        
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setMinimumDate(QDate.currentDate())
        
        # Controles de tempo
        time_group = QWidget()
        time_layout = QHBoxLayout(time_group)
        
        self.hour_spin = QSpinBox()
        self.hour_spin.setRange(0, 23)
        self.hour_spin.setValue(QTime.currentTime().hour())
        
        self.minute_spin = QSpinBox()
        self.minute_spin.setRange(0, 59)
        self.minute_spin.setValue(QTime.currentTime().minute())
        
        self.second_spin = QSpinBox()
        self.second_spin.setRange(0, 59)
        self.second_spin.setValue(QTime.currentTime().second())
        
        time_layout.addWidget(QLabel("Hora:"))
        time_layout.addWidget(self.hour_spin)
        time_layout.addWidget(QLabel("Minuto:"))
        time_layout.addWidget(self.minute_spin)
        time_layout.addWidget(QLabel("Segundo:"))
        time_layout.addWidget(self.second_spin)
        
        calendar_layout.addWidget(self.calendar)
        calendar_layout.addWidget(time_group)
        
        # Grupo do formulário de agendamento
        form_group = QGroupBox("Agendar Cena (Ctrl+A)")
        form_layout = QFormLayout(form_group)
        
        self.scene_combo = QComboBox()
        self.scene_combo.setEditable(False)
        
        self.repeat_check = QCheckBox("Repetir diariamente")
        self.repeat_days = QSpinBox()
        self.repeat_days.setRange(1, 365)
        self.repeat_days.setValue(1)
        self.repeat_days.setEnabled(False)
        
        self.repeat_check.stateChanged.connect(
            lambda: self.repeat_days.setEnabled(self.repeat_check.isChecked())
        )
        
        self.notes_input = QLineEdit()
        self.notes_input.setPlaceholderText("Observações sobre este agendamento")
        
        self.add_schedule_btn = QPushButton("Adicionar Agendamento")
        self.add_schedule_btn.clicked.connect(self._add_schedule)
        self.shortcuts['add_schedule'].activated.connect(self._add_schedule)
        
        form_layout.addRow("Cena:", self.scene_combo)
        form_layout.addRow(self.repeat_check)
        form_layout.addRow("Repetir por (dias):", self.repeat_days)
        form_layout.addRow("Observações:", self.notes_input)
        form_layout.addRow(self.add_schedule_btn)
        
        left_layout.addWidget(calendar_group)
        left_layout.addWidget(form_group)
        
        # Painel direito - Lista de agendamentos
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        schedule_group = QGroupBox("Agendamentos Programados")
        schedule_layout = QVBoxLayout(schedule_group)
        
        self.schedule_table = QTableWidget(0, 6)
        self.schedule_table.setHorizontalHeaderLabels([
            "Data", "Hora", "Cena", "Repetição", "Mídia", "Ações"
        ])
        self.schedule_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.schedule_table.verticalHeader().setVisible(False)
        
        schedule_layout.addWidget(self.schedule_table)
        right_layout.addWidget(schedule_group)
        
        # Adiciona os painéis ao splitter
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([400, 600])
        
        layout.addWidget(splitter)
        self.tabs.addTab(tab, "Agendamento")
        
        # Atualiza a lista de cenas quando a aba é selecionada
        self.tabs.currentChanged.connect(self._on_tab_changed)
        
        # Carrega agendamentos existentes
        self._load_schedules()

    def _setup_reports_tab(self):
        """Configura a aba de relatórios com filtros e visualização."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Controles de filtro
        filter_group = QGroupBox("Filtros")
        filter_layout = QHBoxLayout(filter_group)
        
        # Período
        period_group = QGroupBox("Período")
        period_layout = QHBoxLayout(period_group)
        
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addDays(-7))
        self.start_date.setCalendarPopup(True)
        
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        
        period_layout.addWidget(QLabel("De:"))
        period_layout.addWidget(self.start_date)
        period_layout.addWidget(QLabel("Até:"))
        period_layout.addWidget(self.end_date)
        
        # Filtro por cliente
        self.cliente_filter = QComboBox()
        self.cliente_filter.addItem("Todos os Clientes", None)
        self._load_clientes_combo()
        
        # Botões
        self.generate_btn = QPushButton("Gerar Relatório (Ctrl+G)")
        self.generate_btn.clicked.connect(self._generate_report)
        self.shortcuts['generate_report'].activated.connect(self._generate_report)
        
        self.export_pdf_btn = QPushButton("Exportar para PDF")
        self.export_pdf_btn.clicked.connect(self._export_pdf_report)
        
        filter_layout.addWidget(period_group)
        filter_layout.addWidget(QLabel("Cliente:"))
        filter_layout.addWidget(self.cliente_filter)
        filter_layout.addWidget(self.generate_btn)
        filter_layout.addWidget(self.export_pdf_btn)
        
        # Tabela de resultados
        self.report_table = QTableWidget()
        self.report_table.setColumnCount(8)
        self.report_table.setHorizontalHeaderLabels([
            "Data", "Hora", "Cena", "Tipo", "Status", "Mídia", "Cliente", "Agência"
        ])
        self.report_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.report_table.setSortingEnabled(True)
        
        layout.addWidget(filter_group)
        layout.addWidget(self.report_table)
        self.tabs.addTab(tab, "Relatórios")

    def _setup_cadastro_tab(self):
        """Configura a aba de cadastro com formulários completos."""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Abas de cadastro
        cadastro_tabs = QTabWidget()
        
        # Sub-aba de Emissoras
        emissora_tab = QWidget()
        emissora_layout = QFormLayout(emissora_tab)
        
        self.emissora_nome = QLineEdit()
        self.emissora_razao = QLineEdit()
        
        # Configura máscara para CNPJ
        self.emissora_cnpj = QLineEdit()
        cnpj_regex = QRegExp(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
        self.emissora_cnpj.setValidator(QRegExpValidator(cnpj_regex, self.emissora_cnpj))
        
        self.emissora_endereco = QLineEdit()
        self.emissora_contato = QLineEdit()
        self.emissora_email = QLineEdit()
        
        # Botão de logo com visualização
        self.emissora_logo_btn = QPushButton("Selecionar Logo")
        self.emissora_logo_btn.clicked.connect(self._select_logo)
        self.emissora_logo_preview = QLabel()
        self.emissora_logo_preview.setAlignment(Qt.AlignCenter)
        self.emissora_logo_preview.setFixedSize(100, 100)
        
        self.salvar_emissora_btn = QPushButton("Salvar Emissora")
        self.salvar_emissora_btn.clicked.connect(self._salvar_emissora)
        
        emissora_layout.addRow("Nome Fantasia:", self.emissora_nome)
        emissora_layout.addRow("Razão Social:", self.emissora_razao)
        emissora_layout.addRow("CNPJ (XX.XXX.XXX/XXXX-XX):", self.emissora_cnpj)
        emissora_layout.addRow("Endereço:", self.emissora_endereco)
        emissora_layout.addRow("Contato:", self.emissora_contato)
        emissora_layout.addRow("E-mail:", self.emissora_email)
        emissora_layout.addRow("Logo:", self.emissora_logo_btn)
        emissora_layout.addRow(self.emissora_logo_preview)
        emissora_layout.addRow(self.salvar_emissora_btn)
        
        # Sub-aba de Clientes
        cliente_tab = QWidget()
        cliente_layout = QFormLayout(cliente_tab)
        
        self.cliente_nome = QLineEdit()
        self.cliente_razao = QLineEdit()
        
        # Configura máscara para CNPJ
        self.cliente_cnpj = QLineEdit()
        self.cliente_cnpj.setValidator(QRegExpValidator(cnpj_regex, self.cliente_cnpj))
        
        self.cliente_ie = QLineEdit()
        self.cliente_contato = QLineEdit()
        
        # ComboBox para agência
        self.cliente_agencia = QComboBox()
        self._load_agencias_combo()
        
        self.salvar_cliente_btn = QPushButton("Salvar Cliente")
        self.salvar_cliente_btn.clicked.connect(self._salvar_cliente)
        
        # Tabela de clientes
        self.clientes_table = QTableWidget()
        self.clientes_table.setColumnCount(5)
        self.clientes_table.setHorizontalHeaderLabels(["ID", "Nome", "CNPJ", "Agência", "Ações"])
        self.clientes_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        cliente_layout.addRow("Nome Fantasia:", self.cliente_nome)
        cliente_layout.addRow("Razão Social:", self.cliente_razao)
        cliente_layout.addRow("CNPJ:", self.cliente_cnpj)
        cliente_layout.addRow("Inscrição Estadual:", self.cliente_ie)
        cliente_layout.addRow("Contato:", self.cliente_contato)
        cliente_layout.addRow("Agência:", self.cliente_agencia)
        cliente_layout.addRow(self.salvar_cliente_btn)
        cliente_layout.addRow(self.clientes_table)
        
        # Sub-aba de Agências
        agencia_tab = QWidget()
        agencia_layout = QFormLayout(agencia_tab)
        
        self.agencia_nome = QLineEdit()
        self.agencia_razao = QLineEdit()
        
        # Configura máscara para CNPJ
        self.agencia_cnpj = QLineEdit()
        self.agencia_cnpj.setValidator(QRegExpValidator(cnpj_regex, self.agencia_cnpj))
        
        self.agencia_ie = QLineEdit()
        self.agencia_contato = QLineEdit()
        self.salvar_agencia_btn = QPushButton("Salvar Agência")
        self.salvar_agencia_btn.clicked.connect(self._salvar_agencia)
        
        # Tabela de agências
        self.agencias_table = QTableWidget()
        self.agencias_table.setColumnCount(4)
        self.agencias_table.setHorizontalHeaderLabels(["ID", "Nome", "CNPJ", "Ações"])
        self.agencias_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        agencia_layout.addRow("Nome Fantasia:", self.agencia_nome)
        agencia_layout.addRow("Razão Social:", self.agencia_razao)
        agencia_layout.addRow("CNPJ:", self.agencia_cnpj)
        agencia_layout.addRow("Inscrição Estadual:", self.agencia_ie)
        agencia_layout.addRow("Contato:", self.agencia_contato)
        agencia_layout.addRow(self.salvar_agencia_btn)
        agencia_layout.addRow(self.agencias_table)
        
        # Sub-aba de Mídias
        midia_tab = QWidget()
        midia_layout = QFormLayout(midia_tab)
        
        self.midia_nome = QLineEdit()
        self.midia_cliente = QComboBox()
        self._load_clientes_combo()
        
        self.midia_file_btn = QPushButton("Selecionar Arquivo")
        self.midia_file_btn.clicked.connect(self._select_midia_file)
        self.midia_file_label = QLabel("Nenhum arquivo selecionado")
        
        self.salvar_midia_btn = QPushButton("Salvar Mídia")
        self.salvar_midia_btn.clicked.connect(self._salvar_midia)
        
        # Tabela de mídias
        self.midias_table = QTableWidget()
        self.midias_table.setColumnCount(4)
        self.midias_table.setHorizontalHeaderLabels(["ID", "Nome", "Cliente", "Ações"])
        self.midias_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        midia_layout.addRow("Nome:", self.midia_nome)
        midia_layout.addRow("Cliente:", self.midia_cliente)
        midia_layout.addRow("Arquivo:", self.midia_file_btn)
        midia_layout.addRow(self.midia_file_label)
        midia_layout.addRow(self.salvar_midia_btn)
        midia_layout.addRow(self.midias_table)
        
        # Adiciona as sub-abas
        cadastro_tabs.addTab(emissora_tab, "Emissoras")
        cadastro_tabs.addTab(cliente_tab, "Clientes")
        cadastro_tabs.addTab(agencia_tab, "Agências")
        cadastro_tabs.addTab(midia_tab, "Mídias")
        
        layout.addWidget(cadastro_tabs)
        self.tabs.addTab(tab, "Cadastros")
        
        # Carrega dados iniciais
        self._load_clientes_table()
        self._load_agencias_table()
        self._load_midias_table()

    # ... (implementar os métodos restantes seguindo o mesmo padrão)

def apply_dark_theme(app):
    """Aplica um tema escuro moderno à aplicação."""
    app.setStyle("Fusion")
    
    dark_palette = QPalette()
    dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.WindowText, Qt.white)
    dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
    dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
    dark_palette.setColor(QPalette.ToolTipText, Qt.white)
    dark_palette.setColor(QPalette.Text, Qt.white)
    dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ButtonText, Qt.white)
    dark_palette.setColor(QPalette.BrightText, Qt.red)
    dark_palette.setColor(QPalette.Highlight, QColor(142, 45, 197))
    dark_palette.setColor(QPalette.HighlightedText, Qt.black)
    
    app.setPalette(dark_palette)

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        apply_dark_theme(app)
        
        window = MainWindow()
        if not hasattr(window, 'db') or not window.db.connection:
            raise Exception("Falha na inicialização do banco de dados")
            
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        error_msg = f"Não foi possível iniciar o aplicativo:\n{str(e)}"
        print(error_msg)
        QMessageBox.critical(None, "Erro Fatal", error_msg)
        sys.exit(1)
        
