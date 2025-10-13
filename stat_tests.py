import sys
import os
import logging
from typing import Optional, Tuple, Dict, Any

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QTextEdit,
    QStackedWidget,
    QMessageBox,
    QInputDialog,
)
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from scipy.stats import ttest_ind_from_stats, chi2_contingency
import yaml


# =====================
# Logging
# =====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('stat_tests_gui.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)


# =====================
# Queries (copied/adapted from the provided CLI implementation)
# =====================
TEST_QUERIES: Dict[int, Dict[str, str]] = {
    1: {
        'name': 'Средний чек',
        'sql_male_female': """
            SELECT 
                c.gender AS group_name,
                AVG(t.total_amount) AS mean,
                STDEV(t.total_amount) AS stddev,
                COUNT(*) AS count
            FROM rd_transactions t
            JOIN rd_customers c ON t.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            SELECT 
                s.traffic_source AS group_name,
                AVG(t.total_amount) AS mean,
                STDEV(t.total_amount) AS stddev,
                COUNT(*) AS count
            FROM rd_transactions t
            JOIN rd_sessions s ON t.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY s.traffic_source
        """
    },
    2: {
        'name': 'Количество товаров в заказе',
        'sql_male_female': """
            WITH prods AS (
                SELECT 
                    t.booking_id,
                    t.customer_id,
                    SUM(tp.quantity) AS total_quantity
                FROM rd_transactions t
                JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
                GROUP BY t.booking_id, t.customer_id
            )
            SELECT 
                c.gender AS group_name,
                AVG(CAST(p.total_quantity AS FLOAT)) AS mean,
                STDEV(p.total_quantity) AS stddev,
                COUNT(*) AS count
            FROM prods p
            JOIN rd_customers c ON p.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            WITH prods AS (
                SELECT 
                    t.booking_id,
                    t.session_id,
                    SUM(tp.quantity) AS total_quantity
                FROM rd_transactions t
                JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
                GROUP BY t.booking_id, t.session_id
            )
            SELECT 
                s.traffic_source AS group_name,
                AVG(CAST(p.total_quantity AS FLOAT)) AS mean,
                STDEV(p.total_quantity) AS stddev,
                COUNT(*) AS count
            FROM prods p
            JOIN rd_sessions s ON p.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY s.traffic_source
        """
    },
    3: {
        'name': 'Количество кликов до букинга',
        'sql_male_female': """
            SELECT 
                c.gender AS group_name,
                AVG(CAST(s.CLICK_cnt AS FLOAT)) AS mean,
                STDEV(s.CLICK_cnt) AS stddev,
                COUNT(*) AS count
            FROM rd_sessions s
            JOIN rd_transactions t ON s.session_id = t.session_id
            JOIN rd_customers c ON t.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            SELECT 
                traffic_source AS group_name,
                AVG(CAST(CLICK_cnt AS FLOAT)) AS mean,
                STDEV(CLICK_cnt) AS stddev,
                COUNT(*) AS count
            FROM rd_sessions
            WHERE traffic_source IN ('WEB', 'MOBILE')
            GROUP BY traffic_source
        """
    },
    4: {
        'name': 'Время до покупки (минуты)',
        'sql_male_female': """
            SELECT 
                c.gender AS group_name,
                AVG(DATEDIFF(MINUTE, s.session_start_time, t.created_at)) AS mean,
                STDEV(DATEDIFF(MINUTE, s.session_start_time, t.created_at)) AS stddev,
                COUNT(*) AS count
            FROM rd_transactions t
            JOIN rd_sessions s ON t.session_id = s.session_id
            JOIN rd_customers c ON t.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            SELECT 
                s.traffic_source AS group_name,
                AVG(DATEDIFF(MINUTE, s.session_start_time, t.created_at)) AS mean,
                STDEV(DATEDIFF(MINUTE, s.session_start_time, t.created_at)) AS stddev,
                COUNT(*) AS count
            FROM rd_transactions t
            JOIN rd_sessions s ON t.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY s.traffic_source
        """
    },
    5: {
        'name': 'Доля заказов с мужскими товарами',
        'sql_male_female': """
            WITH mens_prods AS (
                SELECT 
                    t.booking_id,
                    t.customer_id,
                    t.session_id,
                    MAX(CASE WHEN p.gender = 'Men' THEN 1 ELSE 0 END) AS has_mens_prods
                FROM rd_transactions t
                JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
                JOIN rd_prods p ON tp.product_id = p.prod_id
                GROUP BY t.booking_id, t.customer_id, t.session_id
            )
            SELECT 
                c.gender AS group_name,
                AVG(CAST(m.has_mens_prods AS FLOAT)) AS mean,
                STDEV(m.has_mens_prods) AS stddev,
                COUNT(*) AS count
            FROM mens_prods m
            JOIN rd_customers c ON m.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            WITH mens_prods AS (
                SELECT 
                    t.booking_id,
                    t.session_id,
                    MAX(CASE WHEN p.gender = 'Men' THEN 1 ELSE 0 END) AS has_mens_prods
                FROM rd_transactions t
                JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
                JOIN rd_prods p ON tp.product_id = p.prod_id
                GROUP BY t.booking_id, t.session_id
            )
            SELECT 
                s.traffic_source AS group_name,
                AVG(CAST(m.has_mens_prods AS FLOAT)) AS mean,
                STDEV(m.has_mens_prods) AS stddev,
                COUNT(*) AS count
            FROM mens_prods m
            JOIN rd_sessions s ON m.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY s.traffic_source
        """
    },
    6: {
        'name': 'Средняя стоимость доставки',
        'sql_male_female': """
            SELECT 
                c.gender AS group_name,
                AVG(t.shipment_fee) AS mean,
                STDEV(t.shipment_fee) AS stddev,
                COUNT(*) AS count
            FROM rd_transactions t
            JOIN rd_customers c ON t.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            SELECT 
                s.traffic_source AS group_name,
                AVG(t.shipment_fee) AS mean,
                STDEV(t.shipment_fee) AS stddev,
                COUNT(*) AS count
            FROM rd_transactions t
            JOIN rd_sessions s ON t.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY s.traffic_source
        """
    },
    7: {
        'name': 'Количество уникальных товаров',
        'sql_male_female': """
            WITH unique_prods AS (
                SELECT 
                    t.booking_id,
                    t.customer_id,
                    t.session_id,
                    COUNT(DISTINCT tp.product_id) AS unique_prods
                FROM rd_transactions t
                JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
                GROUP BY t.booking_id, t.customer_id, t.session_id
            )
            SELECT 
                c.gender AS group_name,
                AVG(CAST(u.unique_prods AS FLOAT)) AS mean,
                STDEV(u.unique_prods) AS stddev,
                COUNT(*) AS count
            FROM unique_prods u
            JOIN rd_customers c ON u.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY c.gender
        """,
        'sql_web_mobile': """
            WITH unique_prods AS (
                SELECT 
                    t.booking_id,
                    t.session_id,
                    COUNT(DISTINCT tp.product_id) AS unique_prods
                FROM rd_transactions t
                JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
                GROUP BY t.booking_id, t.session_id
            )
            SELECT 
                s.traffic_source AS group_name,
                AVG(CAST(u.unique_prods AS FLOAT)) AS mean,
                STDEV(u.unique_prods) AS stddev,
                COUNT(*) AS count
            FROM unique_prods u
            JOIN rd_sessions s ON u.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY s.traffic_source
        """
    }
}

# For Chi-square test: build 2x2 contingency (has_mens_prods vs group)
CHI2_QUERIES: Dict[str, str] = {
    'male_female': """
        WITH mens_prods AS (
            SELECT 
                t.booking_id,
                c.gender AS group_name,
                MAX(CASE WHEN p.gender = 'Men' THEN 1 ELSE 0 END) AS has_mens_prods
            FROM rd_transactions t
            JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
            JOIN rd_prods p ON tp.product_id = p.prod_id
            JOIN rd_customers c ON t.customer_id = c.customer_id
            WHERE c.gender IN ('M', 'F')
            GROUP BY t.booking_id, c.gender
        )
        SELECT group_name, has_mens_prods, COUNT(*) AS cnt
        FROM mens_prods
        GROUP BY group_name, has_mens_prods
    """,
    'web_mobile': """
        WITH mens_prods AS (
            SELECT 
                t.booking_id,
                s.traffic_source AS group_name,
                MAX(CASE WHEN p.gender = 'Men' THEN 1 ELSE 0 END) AS has_mens_prods
            FROM rd_transactions t
            JOIN rd_transactions_prods tp ON t.booking_id = tp.booking_id
            JOIN rd_prods p ON tp.product_id = p.prod_id
            JOIN rd_sessions s ON t.session_id = s.session_id
            WHERE s.traffic_source IN ('WEB', 'MOBILE')
            GROUP BY t.booking_id, s.traffic_source
        )
        SELECT group_name, has_mens_prods, COUNT(*) AS cnt
        FROM mens_prods
        GROUP BY group_name, has_mens_prods
    """,
}


# =====================
# Helpers: config and DB
# =====================

def get_config_path() -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, 'config.yaml')


def load_db_url_from_config() -> Optional[str]:
    try:
        cfg_path = get_config_path()
        if not os.path.exists(cfg_path):
            return None
        with open(cfg_path, 'r', encoding='utf-8') as f:
            cfg = yaml.safe_load(f) or {}
        return cfg.get('db_url')
    except Exception as e:
        logging.error(f"Failed to read config.yaml: {e}")
        return None


def save_db_url_to_config(db_url: str) -> None:
    try:
        cfg_path = get_config_path()
        with open(cfg_path, 'w', encoding='utf-8') as f:
            yaml.safe_dump({'db_url': db_url}, f, allow_unicode=True, sort_keys=True)
        logging.info("DB URL saved to config.yaml")
    except Exception as e:
        logging.error(f"Failed to save config.yaml: {e}")


def try_create_engine(db_url: str) -> Engine:
    engine = create_engine(db_url)
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    return engine


# =====================
# Stats runners
# =====================

def perform_ttest_from_stats(mean1: float, std1: float, n1: int,
                             mean2: float, std2: float, n2: int,
                             metric_name: str, group1_name: str, group2_name: str) -> Tuple[Optional[float], Optional[float], str]:
    if n1 is None or n2 is None or int(n1) < 2 or int(n2) < 2:
        msg = (f"Предупреждение: Недостаточно данных для {metric_name} между "
               f"{group1_name} и {group2_name} (n1={n1}, n2={n2}).")
        logging.warning(msg)
        return None, None, msg
    t_stat, p_val = ttest_ind_from_stats(mean1, std1, n1, mean2, std2, n2, equal_var=False)
    decision = 'Есть различие' if p_val < 0.05 else 'Нет различий'
    result = (
        f"{metric_name} ({group1_name} vs {group2_name})\n"
        f"t={t_stat:.3f}, p={p_val:.4f} -> {decision}"
    )
    logging.info(result)
    return float(t_stat), float(p_val), result


def run_ttest(engine: Engine, test_choice: int, group_choice: int) -> str:
    try:
        test_info = TEST_QUERIES[test_choice]
        metric_name = test_info['name']
        sql_query = test_info['sql_male_female'] if group_choice == 1 else test_info['sql_web_mobile']
        group1_name, group2_name = ('M', 'F') if group_choice == 1 else ('WEB', 'MOBILE')

        logging.info(f"Executing SQL for t-test: {metric_name}")
        df_stats = pd.read_sql(sql_query, engine)
        if 'group_name' not in df_stats.columns:
            raise ValueError("SQL must return 'group_name' column")

        if len(df_stats) < 2:
            found = df_stats.get('group_name', pd.Series(dtype=str)).tolist()
            return (
                f"Ошибка: Для t-теста нужно 2 группы. Найдено {len(df_stats)}.\n"
                f"Группы: {found}"
            )

        group1_stats = df_stats[df_stats['group_name'].astype(str).str.contains(group1_name, case=False)]
        group2_stats = df_stats[df_stats['group_name'].astype(str).str.contains(group2_name, case=False)]
        if group1_stats.empty or group2_stats.empty:
            return f"Ошибка: Нет данных для {group1_name} или {group2_name}."

        mean1, std1, n1 = group1_stats.iloc[0][['mean', 'stddev', 'count']]
        mean2, std2, n2 = group2_stats.iloc[0][['mean', 'stddev', 'count']]

        t_stat, p_val, text = perform_ttest_from_stats(mean1, std1, n1, mean2, std2, n2, metric_name, group1_name, group2_name)
        if p_val is None:
            return text
        details = (
            f"{metric_name}\n"
            f"{group1_name}: mean={float(mean1):.3f}, std={float(std1):.3f}, n={int(n1)}\n"
            f"{group2_name}: mean={float(mean2):.3f}, std={float(std2):.3f}, n={int(n2)}\n"
            f"{text}"
        )
        return details

    except Exception as e:
        logging.error(f"T-test failed: {e}")
        return f"Ошибка: {e}"


def run_chi_square(engine: Engine, group_choice: int) -> str:
    try:
        key = 'male_female' if group_choice == 1 else 'web_mobile'
        sql = CHI2_QUERIES[key]
        logging.info("Executing SQL for Chi-square (mens products proportion)")
        df = pd.read_sql(sql, engine)
        if df.empty:
            return "Ошибка: Нет данных для хи-квадрат теста."
        # Build contingency table: rows=group_name, cols=has_mens_prods (0/1)
        cont = pd.pivot_table(df, index='group_name', columns='has_mens_prods', values='cnt', aggfunc='sum', fill_value=0)
        # Ensure both columns 0 and 1 exist
        for col in [0, 1]:
            if col not in cont.columns:
                cont[col] = 0
        cont = cont[[0, 1]]
        chi2, p_value, dof, expected = chi2_contingency(cont)
        g1, g2 = cont.index.tolist()[:2] if len(cont.index) >= 2 else ("Group1", "Group2")
        res_lines = [
            "Доля заказов с мужскими товарами",
            f"Группы: {', '.join(cont.index.astype(str))}",
            f"Таблица сопряженности (0,1):\n{cont.to_string()}",
            f"Chi2={chi2:.3f}, p={p_value:.4f} -> {'Есть различие' if p_value < 0.05 else 'Нет различий'}",
        ]
        return "\n".join(res_lines)
    except Exception as e:
        logging.error(f"Chi-square failed: {e}")
        return f"Ошибка: {e}"


# =====================
# GUI
# =====================
class StatsTestsWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Статистические тесты")
        self.resize(900, 600)

        self.engine: Optional[Engine] = None
        self.selected_test: str = 'chi2'  # 'chi2' or 'ttest'
        self.selected_metric: Optional[int] = None
        self.selected_groups: Optional[int] = None  # 1 or 2

        central = QWidget(self)
        root = QVBoxLayout(central)
        self.setCentralWidget(central)

        # Top status and actions
        top = QHBoxLayout()
        self.db_status = QLabel("DB: не подключено")
        self.db_status.setStyleSheet("color: #a33; font-size: 12px;")
        top.addWidget(self.db_status)
        top.addStretch(1)
        self.btn_config = QPushButton("Подключение…")
        self.btn_config.clicked.connect(self.on_setup_connection)
        top.addWidget(self.btn_config)
        root.addLayout(top)

        # Pages
        self.pages = QStackedWidget(self)
        self.page_test = self._build_page_test()
        self.page_metric = self._build_page_metric()
        self.page_groups = self._build_page_groups()
        self.page_result = self._build_page_result()
        self.pages.addWidget(self.page_test)
        self.pages.addWidget(self.page_metric)
        self.pages.addWidget(self.page_groups)
        self.pages.addWidget(self.page_result)
        root.addWidget(self.pages, stretch=1)

        self._update_db_status()

    # ----- Pages builders
    def _build_page_test(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)

        lbl = QLabel("Выберите тест:")
        lbl.setStyleSheet("font-size: 16px;")
        layout.addWidget(lbl)

        self.test_combo = QComboBox()
        self.test_combo.addItems(["Хи-квадрат тест", "t-тест"])
        self.test_combo.setCurrentIndex(0)
        self.test_combo.currentIndexChanged.connect(self.on_test_changed)
        layout.addWidget(self.test_combo)

        layout.addStretch(1)

        nav = QHBoxLayout()
        nav.addStretch(1)
        btn_next = QPushButton("Далее")
        btn_next.clicked.connect(self.on_next_from_test)
        nav.addWidget(btn_next)
        layout.addLayout(nav)

        return page

    def _build_page_metric(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)

        lbl = QLabel("Выберите вопрос тестирования:")
        lbl.setStyleSheet("font-size: 16px;")
        layout.addWidget(lbl)

        self.metric_combo = QComboBox()
        self.metric_combo.addItems([
            TEST_QUERIES[i]['name'] for i in sorted(TEST_QUERIES.keys())
        ])
        layout.addWidget(self.metric_combo)

        layout.addStretch(1)

        nav = QHBoxLayout()
        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(lambda: self.pages.setCurrentWidget(self.page_test))
        nav.addWidget(btn_back)
        nav.addStretch(1)
        btn_next = QPushButton("Далее")
        btn_next.clicked.connect(self.on_next_from_metric)
        nav.addWidget(btn_next)
        layout.addLayout(nav)

        return page

    def _build_page_groups(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)

        lbl = QLabel("Выберите группы для сравнения:")
        lbl.setStyleSheet("font-size: 16px;")
        layout.addWidget(lbl)

        self.groups_combo = QComboBox()
        self.groups_combo.addItems(["Мужчины vs Женщины", "WEB vs MOBILE"])
        layout.addWidget(self.groups_combo)

        layout.addStretch(1)

        nav = QHBoxLayout()
        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.on_back_from_groups)
        nav.addWidget(btn_back)
        nav.addStretch(1)
        btn_run = QPushButton("Запустить")
        btn_run.clicked.connect(self.on_run_test)
        nav.addWidget(btn_run)
        layout.addLayout(nav)

        return page

    def _build_page_result(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setStyleSheet("font-size: 14px;")
        layout.addWidget(self.result_text, stretch=1)

        nav = QHBoxLayout()
        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.on_back_from_result)
        nav.addWidget(btn_back)
        nav.addStretch(1)
        btn_new = QPushButton("Новый тест")
        btn_new.clicked.connect(self.on_new_test)
        nav.addWidget(btn_new)
        btn_exit = QPushButton("Выход")
        btn_exit.clicked.connect(self.close)
        nav.addWidget(btn_exit)
        layout.addLayout(nav)

        return page

    # ----- Navigation handlers
    def on_test_changed(self, idx: int) -> None:
        self.selected_test = 'chi2' if idx == 0 else 'ttest'

    def on_next_from_test(self) -> None:
        if self.selected_test == 'ttest':
            self.pages.setCurrentWidget(self.page_metric)
        else:
            self.pages.setCurrentWidget(self.page_groups)

    def on_next_from_metric(self) -> None:
        self.selected_metric = self.metric_combo.currentIndex() + 1  # 1..7
        self.pages.setCurrentWidget(self.page_groups)

    def on_back_from_groups(self) -> None:
        if self.selected_test == 'ttest':
            self.pages.setCurrentWidget(self.page_metric)
        else:
            self.pages.setCurrentWidget(self.page_test)

    def on_back_from_result(self) -> None:
        self.pages.setCurrentWidget(self.page_groups)

    def on_new_test(self) -> None:
        self.selected_test = 'chi2'
        self.test_combo.setCurrentIndex(0)
        self.metric_combo.setCurrentIndex(0)
        self.groups_combo.setCurrentIndex(0)
        self.pages.setCurrentWidget(self.page_test)

    # ----- Connection setup
    def _update_db_status(self) -> None:
        ok = self.engine is not None
        self.db_status.setText("DB: подключено" if ok else "DB: не подключено")
        self.db_status.setStyleSheet("color: #2a7; font-size: 12px;" if ok else "color: #a33; font-size: 12px;")

    def ensure_connection(self) -> bool:
        if self.engine is not None:
            return True
        # Try config
        db_url = load_db_url_from_config()
        if db_url:
            try:
                self.engine = try_create_engine(db_url)
                self._update_db_status()
                return True
            except Exception as e:
                logging.error(f"Config DB connection failed: {e}")
        # Ask user
        db_url, ok = QInputDialog.getText(self, "Подключение к БД", "Введите SQLAlchemy DB URL:")
        if not ok or not db_url:
            QMessageBox.warning(self, "Подключение", "Подключение к БД не установлено.")
            self._update_db_status()
            return False
        try:
            self.engine = try_create_engine(db_url)
            self._update_db_status()
            # Ask to save
            save = QMessageBox.question(self, "Сохранить конфиг", "Сохранить DB URL в config.yaml?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if save == QMessageBox.Yes:
                save_db_url_to_config(db_url)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Ошибка подключения", f"Не удалось подключиться: {e}")
            self.engine = None
            self._update_db_status()
            return False

    # ----- Run tests
    def on_run_test(self) -> None:
        self.selected_groups = 1 if self.groups_combo.currentIndex() == 0 else 2
        if not self.ensure_connection():
            return
        try:
            if self.selected_test == 'chi2':
                text = run_chi_square(self.engine, self.selected_groups)
            else:
                if not self.selected_metric:
                    QMessageBox.warning(self, "T-тест", "Выберите метрику перед запуском.")
                    return
                text = run_ttest(self.engine, self.selected_metric, self.selected_groups)
            self.result_text.setPlainText(text)
            self.pages.setCurrentWidget(self.page_result)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось выполнить тест: {e}")

    # ----- Manual connection setup
    def on_setup_connection(self) -> None:
        # Force reconfigure regardless of current state
        self.engine = None
        self._update_db_status()
        self.ensure_connection()


# =====================
# Entrypoint
# =====================

def run_gui() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    window = StatsTestsWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    run_gui()
