import sys
import os
from typing import Optional

import pandas as pd

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QUrl
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QStackedWidget,
    QLabel,
    QPushButton,
    QComboBox,
    QTextEdit,
    QTableView,
    QSizePolicy,
    QHeaderView,
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas

from utils import (
    customers,
    columns_str,
    compute_ltv_factors_for_column,
    compute_ltv_cohort_for_column,
    compute_revenue_structure_for_column,
    compute_chi2_result,
    compute_ttest_result,
    create_bar_plot,
    create_line_plot,
    create_pie_plot,
)


class PandasModel(QAbstractTableModel):
    def __init__(self, df: Optional[pd.DataFrame] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._df = df.copy() if df is not None else pd.DataFrame()

    def setDataFrame(self, df: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = df.copy() if df is not None else pd.DataFrame()
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if self._df is None else self._df.shape[0]

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if self._df is None else self._df.shape[1]

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid() or role not in (Qt.DisplayRole, Qt.EditRole):
            return None
        value = self._df.iat[index.row(), index.column()]
        if pd.isna(value):
            return ""
        return str(value)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            try:
                return str(self._df.columns[section])
            except Exception:
                return ""
        else:
            try:
                return str(self._df.index[section])
            except Exception:
                return ""


class PlotWidget(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._canvas: Optional[FigureCanvas] = None

    def set_figure(self, fig) -> None:
        if self._canvas is not None:
            self._layout.removeWidget(self._canvas)
            self._canvas.setParent(None)
            self._canvas.deleteLater()
            self._canvas = None
        if fig is not None:
            self._canvas = FigureCanvas(fig)
            self._canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self._layout.addWidget(self._canvas)
            self._canvas.draw()


class SummaryPage(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        title = QLabel("Executive Summary")
        title.setStyleSheet("font-weight: bold; font-size: 20px;")
        layout.addWidget(title)

        self.message = QLabel("")
        self.message.setStyleSheet("font-size: 14px; color: #666;")
        self.message.setVisible(False)
        layout.addWidget(self.message)

        self.web = QWebEngineView(self)
        self.web.settings().setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)
        layout.addWidget(self.web)

        self._pdf_doc = QPdfDocument(self)
        self._pdf_view = QPdfView(self)
        self._pdf_view.setVisible(False)
        layout.addWidget(self._pdf_view)

        self.load_summary()

    def load_summary(self) -> None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        pdf_path = os.path.join(base_dir, 'Executive_summary.pdf')
        if not os.path.exists(pdf_path):
            self.web.setVisible(False)
            self._pdf_view.setVisible(False)
            self.message.setVisible(True)
            self.message.setText(f"Executive_summary.pdf not found at: {pdf_path}")
            return

        def on_loaded(ok: bool) -> None:
            if ok:
                self.message.setVisible(False)
                self._pdf_view.setVisible(False)
                self.web.setVisible(True)
            else:
                self.web.setVisible(False)
                try:
                    self._pdf_doc.load(pdf_path)
                    self._pdf_view.setDocument(self._pdf_doc)
                    self._pdf_view.setVisible(True)
                    self.message.setVisible(False)
                except Exception as e:
                    self._pdf_view.setVisible(False)
                    self.message.setVisible(True)
                    self.message.setText(f"Unable to open PDF: {e}")

        self.web.loadFinished.connect(on_loaded)
        self.web.setUrl(QUrl.fromLocalFile(pdf_path))


class LtvFactorsPage(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.key_by_label = {v: k for k, v in columns_str.items()}

        root = QVBoxLayout(self)

        controls = QHBoxLayout()
        lbl = QLabel("Split by:")
        lbl.setStyleSheet("font-size: 14px;")
        controls.addWidget(lbl)
        self.combo = QComboBox()
        self.combo.addItems(list(self.key_by_label.keys()))
        self.combo.setStyleSheet("font-size: 14px;")
        controls.addWidget(self.combo)
        controls.addStretch(1)
        self.run_btn = QPushButton("Build")
        self.run_btn.setStyleSheet("font-size: 14px;")
        self.run_btn.clicked.connect(self.on_build)
        controls.addWidget(self.run_btn)
        root.addLayout(controls)

        self.table = QTableView(self)
        self.table_model = PandasModel()
        self.table.setModel(self.table_model)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet("font-size: 14px;")
        root.addWidget(self.table, stretch=1)

        self.plot = PlotWidget(self)
        root.addWidget(self.plot, stretch=2)

        self.on_build()

    def on_build(self) -> None:
        try:
            key = self.key_by_label.get(self.combo.currentText())
            metrics, title, formatters = compute_ltv_factors_for_column(customers, key)
            self.table_model.setDataFrame(metrics)
            fig = create_bar_plot(metrics, title, formatters, figsize=(16, 6), show=False)
            self.plot.set_figure(fig)
        except Exception as e:
            self.plot.set_figure(None)
            self.table_model.setDataFrame(pd.DataFrame())


class LtvCohortsPage(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.key_by_label = {v: k for k, v in columns_str.items()}

        root = QVBoxLayout(self)

        controls = QHBoxLayout()
        lbl = QLabel("Split by:")
        lbl.setStyleSheet("font-size: 14px;")
        controls.addWidget(lbl)
        self.combo = QComboBox()
        self.combo.addItems(list(self.key_by_label.keys()))
        self.combo.setStyleSheet("font-size: 14px;")
        controls.addWidget(self.combo)
        controls.addStretch(1)
        self.run_btn = QPushButton("Build")
        self.run_btn.setStyleSheet("font-size: 14px;")
        self.run_btn.clicked.connect(self.on_build)
        controls.addWidget(self.run_btn)
        root.addLayout(controls)

        self.plot = PlotWidget(self)
        root.addWidget(self.plot, stretch=1)

        self.on_build()

    def on_build(self) -> None:
        try:
            key = self.key_by_label.get(self.combo.currentText())
            metric_ltv, metric_returned_cust, title, index_name = compute_ltv_cohort_for_column(customers, key)
            fig = create_line_plot(metric_ltv, metric_returned_cust, title, index_name, figsize=(16, 9), show=False)
            self.plot.set_figure(fig)
        except Exception as e:
            self.plot.set_figure(None)


class RevenueStructurePage(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.key_by_label = {v: k for k, v in columns_str.items()}

        root = QVBoxLayout(self)

        controls = QHBoxLayout()
        lbl = QLabel("Split by:")
        lbl.setStyleSheet("font-size: 14px;")
        controls.addWidget(lbl)
        self.combo = QComboBox()
        self.combo.addItems(list(self.key_by_label.keys()))
        self.combo.setStyleSheet("font-size: 14px;")
        controls.addWidget(self.combo)
        controls.addStretch(1)
        self.run_btn = QPushButton("Build")
        self.run_btn.setStyleSheet("font-size: 14px;")
        self.run_btn.clicked.connect(self.on_build)
        controls.addWidget(self.run_btn)
        root.addLayout(controls)

        self.table = QTableView(self)
        self.table_model = PandasModel()
        self.table.setModel(self.table_model)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet("font-size: 14px;")
        root.addWidget(self.table, stretch=1)

        self.plot = PlotWidget(self)
        root.addWidget(self.plot, stretch=2)

        self.on_build()

    def on_build(self) -> None:
        try:
            key = self.key_by_label.get(self.combo.currentText())
            metrics, title = compute_revenue_structure_for_column(customers, key)
            self.table_model.setDataFrame(metrics)
            fig = create_pie_plot(metrics, title, figsize=(14, 6), show=False)
            self.plot.set_figure(fig)
        except Exception as e:
            self.plot.set_figure(None)
            self.table_model.setDataFrame(pd.DataFrame())


class StatsTestsPage(QWidget):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        root = QVBoxLayout(self)

        controls = QHBoxLayout()
        lbl = QLabel("Test:")
        lbl.setStyleSheet("font-size: 14px;")
        controls.addWidget(lbl)
        self.test_combo = QComboBox()
        self.test_combo.addItems(["Chi-square test", "T-test"])
        self.test_combo.setStyleSheet("font-size: 14px;")
        self.test_combo.currentIndexChanged.connect(self.on_test_change)
        controls.addWidget(self.test_combo)
        controls.addStretch(1)
        root.addLayout(controls)

        self.stack = QStackedWidget(self)
        self.page_chi2 = self._build_chi2_page()
        self.page_ttest = self._build_ttest_page()
        self.stack.addWidget(self.page_chi2)
        self.stack.addWidget(self.page_ttest)
        root.addWidget(self.stack, stretch=1)

        self.on_test_change(0)

    def _build_chi2_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)

        ctrl = QHBoxLayout()
        lbl = QLabel("Across:")
        lbl.setStyleSheet("font-size: 14px;")
        ctrl.addWidget(lbl)
        self.chi_across = QComboBox()
        self.chi_mapping = {
            'By countries': ('customer_country', 'Customer Country'),
            'By payment methods': ('first_payment_method', 'First purchase payment method'),
        }
        self.chi_across.addItems(list(self.chi_mapping.keys()))
        self.chi_across.setStyleSheet("font-size: 14px;")
        ctrl.addWidget(self.chi_across)
        ctrl.addStretch(1)
        self.chi_run = QPushButton("Run")
        self.chi_run.setStyleSheet("font-size: 14px;")
        self.chi_run.clicked.connect(self.on_run_chi2)
        ctrl.addWidget(self.chi_run)
        layout.addLayout(ctrl)

        self.table_chi_counts = QTableView(self)
        self.model_chi_counts = PandasModel()
        self.table_chi_counts.setModel(self.model_chi_counts)
        self.table_chi_counts.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_chi_counts.setStyleSheet("font-size: 14px;")
        label_counts = QLabel("Number of customers:")
        label_counts.setStyleSheet("font-size: 14px;")
        layout.addWidget(label_counts)
        layout.addWidget(self.table_chi_counts)

        self.table_chi_percent = QTableView(self)
        self.model_chi_percent = PandasModel()
        self.table_chi_percent.setModel(self.model_chi_percent)
        self.table_chi_percent.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_chi_percent.setStyleSheet("font-size: 14px;")
        label_percent = QLabel("Number of customers. % of totals by selected dimension:")
        label_percent.setStyleSheet("font-size: 14px;")
        layout.addWidget(label_percent)
        layout.addWidget(self.table_chi_percent)

        self.chi_text = QTextEdit(self)
        self.chi_text.setReadOnly(True)
        self.chi_text.setStyleSheet("font-size: 16px;")
        layout.addWidget(self.chi_text)

        return page

    def _build_ttest_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)

        ctrl = QHBoxLayout()
        lbl1 = QLabel("Country 1:")
        lbl1.setStyleSheet("font-size: 14px;")
        ctrl.addWidget(lbl1)
        self.ttest_c1 = QComboBox()
        self.ttest_c1.setStyleSheet("font-size: 14px;")
        ctrl.addWidget(self.ttest_c1)
        lbl2 = QLabel("Country 2:")
        lbl2.setStyleSheet("font-size: 14px;")
        ctrl.addWidget(lbl2)
        self.ttest_c2 = QComboBox()
        self.ttest_c2.setStyleSheet("font-size: 14px;")
        ctrl.addWidget(self.ttest_c2)
        ctrl.addStretch(1)
        self.ttest_run = QPushButton("Run")
        self.ttest_run.setStyleSheet("font-size: 14px;")
        self.ttest_run.clicked.connect(self.on_run_ttest)
        ctrl.addWidget(self.ttest_run)
        layout.addLayout(ctrl)

        # Populate countries
        try:
            countries = sorted([c for c in customers['customer_country'].dropna().unique()])
        except Exception:
            countries = []
        self.ttest_c1.addItems(countries)
        self.ttest_c2.addItems(countries)

        self.table_t_counts = QTableView(self)
        self.model_t_counts = PandasModel()
        self.table_t_counts.setModel(self.model_t_counts)
        self.table_t_counts.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_t_counts.setStyleSheet("font-size: 14px;")
        lbl_counts = QLabel("Number of customers:")
        lbl_counts.setStyleSheet("font-size: 14px;")
        layout.addWidget(lbl_counts)
        layout.addWidget(self.table_t_counts)

        self.table_t_percent = QTableView(self)
        self.model_t_percent = PandasModel()
        self.table_t_percent.setModel(self.model_t_percent)
        self.table_t_percent.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_t_percent.setStyleSheet("font-size: 14px;")
        lbl_pct = QLabel("Percentage of Returned customer per selected countries:")
        lbl_pct.setStyleSheet("font-size: 14px;")
        layout.addWidget(lbl_pct)
        layout.addWidget(self.table_t_percent)

        self.t_text = QTextEdit(self)
        self.t_text.setReadOnly(True)
        self.t_text.setStyleSheet("font-size: 16px;")
        layout.addWidget(self.t_text)

        return page

    def on_test_change(self, index: int) -> None:
        self.stack.setCurrentIndex(index)

    def on_run_chi2(self) -> None:
        try:
            across_label = self.chi_across.currentText()
            col_key, col_label = self.chi_mapping[across_label]
            res = compute_chi2_result(customers, 'returned_customer', 'Returned customer', col_key, col_label)
            self.model_chi_counts.setDataFrame(res['contingency_table'])
            self.model_chi_percent.setDataFrame(res['contingency_table_percent'])
            msg = []
            msg.append(f"Does percentage of Returned customer differ across {col_label}?")
            msg.append(res['null_hypothesis'])
            msg.append(f"P-value of Chi-square test = {res['p_value']}")
            if res['decision'] == 'reject':
                msg.append('We reject the null hypothesis.')
            else:
                msg.append('We fail to reject the null hypothesis.')
            msg.append(res['interpretation'])
            self.chi_text.setPlainText("\n".join(msg))
        except Exception as e:
            self.chi_text.setPlainText(f"Error running Chi-square test: {e}")
            self.model_chi_counts.setDataFrame(pd.DataFrame())
            self.model_chi_percent.setDataFrame(pd.DataFrame())

    def on_run_ttest(self) -> None:
        try:
            g1 = self.ttest_c1.currentText()
            g2 = self.ttest_c2.currentText()
            if not g1 or not g2 or g1 == g2:
                self.t_text.setPlainText("Please select two different countries.")
                return
            res = compute_ttest_result(
                customers,
                'returned_customer',
                'Returned customer',
                'customer_country',
                'Customer Country',
                g1,
                g2,
            )
            self.model_t_counts.setDataFrame(res['contingency_table'])
            # percent_true is a Series
            percent_df = pd.DataFrame({"Returned %": res['percent_true'].round(2)})
            self.model_t_percent.setDataFrame(percent_df)
            msg = []
            msg.append(
                f"Is there a significant difference between percentage of Returned customer for {g1} and {g2}?"
            )
            msg.append(res['null_hypothesis'])
            # Safe access in case key missing
            try:
                p1 = float(res['percent_true'].get(g1, float('nan')))
                p2 = float(res['percent_true'].get(g2, float('nan')))
            except Exception:
                p1 = p2 = float('nan')
            msg.append(f"Percentage of Returned customer for {g1} = {round(p1, 2)}")
            msg.append(f"Percentage of Returned customer for {g2} = {round(p2, 2)}")
            msg.append(f"P-value of t-test (independent samples) = {res['p_value']}")
            if res['decision'] == 'reject':
                msg.append('We reject the null hypothesis.')
            else:
                msg.append('We fail to reject the null hypothesis.')
            msg.append(res['interpretation'])
            self.t_text.setPlainText("\n".join(msg))
        except Exception as e:
            self.t_text.setPlainText(f"Error running T-test: {e}")
            self.model_t_counts.setDataFrame(pd.DataFrame())
            self.model_t_percent.setDataFrame(pd.DataFrame())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Global Fashion Retail Sales")
        self.resize(1280, 800)

        central = QWidget(self)
        layout = QHBoxLayout(central)
        self.setCentralWidget(central)

        # Sidebar
        self.sidebar = QListWidget(self)
        self.sidebar.setFixedWidth(260)
        self.sidebar.setStyleSheet("font-size: 16px;")
        self.sidebar.addItem(QListWidgetItem("Executive Summary"))
        self.sidebar.addItem(QListWidgetItem("LTV Factors"))
        self.sidebar.addItem(QListWidgetItem("LTV Cohorts"))
        self.sidebar.addItem(QListWidgetItem("Revenue Structure"))
        self.sidebar.addItem(QListWidgetItem("Statistical Tests"))
        layout.addWidget(self.sidebar)

        # Pages
        self.pages = QStackedWidget(self)
        self.page_summary = SummaryPage(self)
        self.page_ltv_factors = LtvFactorsPage(self)
        self.page_ltv_cohorts = LtvCohortsPage(self)
        self.page_revenue = RevenueStructurePage(self)
        self.page_stats = StatsTestsPage(self)
        self.pages.addWidget(self.page_summary)
        self.pages.addWidget(self.page_ltv_factors)
        self.pages.addWidget(self.page_ltv_cohorts)
        self.pages.addWidget(self.page_revenue)
        self.pages.addWidget(self.page_stats)
        layout.addWidget(self.pages, stretch=1)

        self.sidebar.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.sidebar.setCurrentRow(0)


def run_gui() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()