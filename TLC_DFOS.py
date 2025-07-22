import sys
import os
import openpyxl
import shutil
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from matplotlib.patches import FancyArrowPatch
from matplotlib.ticker import FuncFormatter
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QSizePolicy
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QStackedWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QLineEdit, QInputDialog, QMessageBox,
    QApplication, QTableWidget, QTableWidgetItem
)
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtGui import QPixmap, QDesktopServices
from PyQt5.QtCore import Qt, QUrl

class AspectRatioLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._pixmap = None

    def setPixmap(self, pixmap):
        self._pixmap = pixmap
        if not self.size().isEmpty():
            scaled = self._pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            super().setPixmap(scaled)

    def resizeEvent(self, event):
        if self._pixmap:
            scaled = self._pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            super().setPixmap(scaled)
        super().resizeEvent(event)

def read_excel(excel_path, parent=None):
    """
    Liest eine Excel-Datei ein, zeigt einen Info-Dialog und setzt den Mauszeiger auf 'busy'.
    """
    from PyQt5.QtWidgets import QDialog, QLabel, QVBoxLayout, QMessageBox, QApplication
    from PyQt5.QtCore import Qt
    from PyQt5.QtGui import QCursor
    import pandas as pd
    import os

    if not os.path.exists(excel_path):
        QMessageBox.critical(parent, "Error", f"Datei nicht gefunden:\n{excel_path}")
        return None

    class InfoDialog(QDialog):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("Reading File")
            layout = QVBoxLayout(self)
            label = QLabel("Loading Excel file…\nThis may take a few moments for large files.")
            label.setAlignment(Qt.AlignCenter)
            layout.addWidget(label)
            self.setModal(True)
            self.setFixedSize(580, 220)

    # Set Busy Cursor
    QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))

    info_dialog = InfoDialog(parent)
    info_dialog.show()
    QApplication.processEvents()

    df = None
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        info_dialog.close()
        QApplication.restoreOverrideCursor()
        QMessageBox.critical(parent, "Error", f"Fehler beim Lesen der Datei:\n{e}")
        return None

    info_dialog.close()
    QApplication.restoreOverrideCursor()
    return df






def plot_integral_with_max(df, output_folder, filename="integral_plot.pdf", progress_callback=None):
    r"""
    Erstellt einen Plot mit:
      - Liniendiagramm (Zeit vs. Integral der Deformation)
      - Roter Linie an der Stelle des maximalen Integrals
      - y-Achsen-Label = r'$\int_{0}^{L} \varepsilon \,\mathrm{d}x$ [-‰]'
      - Gitternetz (Major/Minor) und angepasste Tick-Labels

    Die Funktion iteriert über alle Zeilen (df_clean) und ruft bei jeder Iteration
    den progress_callback(current, total) auf, sofern vorhanden.

    Speichert den Plot als PDF und PNG in output_folder und gibt (fig, max_time) zurück.
    """
    os.makedirs(output_folder, exist_ok=True)

    time_column = df.columns[-1]
    deformation_columns = df.columns[:-1]

    # Zeilen entfernen, in denen ALLE Deformationswerte NaN sind
    df_clean = df.dropna(subset=deformation_columns, how='all')
    if df_clean.empty:
        print("Plot wird nicht erstellt, da alle Zeilen NaN sind.")
        return None, None

    integrals = []
    times = []
    n = len(df_clean)

    for i, (_, row) in enumerate(df_clean.iterrows()):
        # --- WICHTIGE ÄNDERUNG: partielle NaNs durch 0 ersetzen ---
        # (Alternativ könnte man hier auch interpolieren)
        deformation_values = row[deformation_columns].fillna(0)

        # Integral berechnen
        val = np.trapz(deformation_values)
        integrals.append(val)

        # Zeitwert erfassen
        times.append(row[time_column])

        # Fortschrittsanzeige aktualisieren, falls gewünscht
        if progress_callback is not None:
            progress_callback(i + 1, n)

    if not integrals:
        print("Keine Integrale berechnet, Plot wird nicht erstellt.")
        return None, None

    max_idx = np.argmax(integrals)
    max_time = times[max_idx]

    # Erstelle den Plot
    fig, ax_line = plt.subplots(figsize=(8, 6))
    ax_line.plot(times, integrals, color='black', linestyle='-', linewidth=1.5, label='Integral Deformation')
    ax_line.axvline(x=max_time, color='red', linestyle='--', linewidth=4.0, label='Zeit vor Riss')
    ax_line.set_xlabel(r'$t \ [\mathrm{s}]$', fontsize=12)
    ax_line.set_ylabel(r'$\int \varepsilon \,\mathrm{d}x$ [-‰]', fontsize=12)
    ax_line.tick_params(axis='both', which='major', labelsize=12)
    ax_line.grid(True, which='major', linestyle='--', linewidth=1, zorder=0)
    ax_line.grid(True, which='minor', linestyle=':', linewidth=0.5, zorder=0)
    ax_line.minorticks_on()
    plt.tight_layout()

    basename = os.path.splitext(filename)[0]
    pdf_path = os.path.join(output_folder, f"{basename}.pdf")
    png_path = os.path.join(output_folder, f"{basename}.png")
    plt.savefig(pdf_path, format='pdf')
    plt.savefig(png_path, format='png')
    plt.close()

    return fig, max_time



def plot_results(result_list, live_end, dead_end, l_ol, eps, max_bin_edges,
                 output_folder, file, min_time, y_limits=None, figsize=(10, 6)):
    x_vals = [pt[1] for pt in result_list]
    y_vals = [pt[0] for pt in result_list]

    fig = plt.figure(figsize=figsize)
    gs = GridSpec(1, 2, width_ratios=[1, 3], wspace=0.04)
    ax_hist = fig.add_subplot(gs[0, 0])
    ax = fig.add_subplot(gs[0, 1], sharey=ax_hist)

    # Histogramm
    if eps and eps > 0:
        bins = np.arange(0, np.nanmax(y_vals) + eps, eps)
    else:
        bins = 10
    counts, bin_edges = np.histogram(y_vals, bins=bins)
    y_pos = (bin_edges[:-1] + bin_edges[1:]) / 2
    max_bin = np.argmax(counts)
    height = bin_edges[1] - bin_edges[0]
    for i, cnt in enumerate(counts):
        color = 'lightblue' if i == max_bin else 'gray'
        ax_hist.barh(y_pos[i], -cnt, height=height, color=color, edgecolor='black')

    ax_hist.set_xlabel(r'$n$', fontsize=14)
    ax_hist.set_ylabel(r'$\epsilon_\mathrm{c}\ [-‰]$', fontsize=14, labelpad=14)
    ax_hist.tick_params(axis="y", left=True, right=False, labelleft=True)
    ax_hist.yaxis.set_label_position("left")
    ax_hist.yaxis.set_ticks_position('left')
    ax_hist.grid(axis='y', ls='--', lw=0.5)
    ax_hist.set_xlim(-max(counts) * 1.1, 0)

    # Rechte Achse: keine Y-Ticks, keine Y-Label
    ax.plot(x_vals, y_vals, 'k-', lw=1.5, label='DFOS')
    if y_limits:
        ax.set_ylim(y_limits)
    if live_end is not None:
        ax.axvline(live_end, color='#13338E', ls='--', lw=1.5)
    if dead_end is not None:
        ax.axvline(x_vals[-1] - dead_end, color='#13338E', ls='--', lw=1.5)
    if max_bin_edges:
        ax.axhspan(max_bin_edges[0], max_bin_edges[1], color='lightblue', alpha=0.7, label='RMS')
    ax.plot([], [], ' ', label=rf'$l_{{ol}}={l_ol:.1f}\,\mathrm{{mm}},\Delta \epsilon_{{c}}={eps:.3f}\,\mathrm{{‰}}$')
    ax.set_xlabel(r'$x\ [\mathrm{mm}]$', fontsize=14)

    # Hier alles y-bezogene ausschalten:
    ax.set_ylabel("")
    ax.tick_params(axis="y", left=False, right=False, labelleft=False, labelright=False)
    ax.grid(True, which='both', ls='--', lw=0.5)

    # Y-Limits synchronisieren (falls nötig)
    ymin, ymax = ax_hist.get_ylim()
    ax.set_ylim(ymin, ymax)

    # Pfeile für Live End und Dead End
    y0, y1 = ax.get_ylim()
    y_arrow = y0 + (y1 - y0) / 3
    x0, x1 = ax.get_xlim()
    if live_end is not None:
        le = np.clip(live_end, x0, x1)
        arrow = FancyArrowPatch((0, y_arrow), (le, y_arrow),
                                arrowstyle='<|-|>', mutation_scale=20,
                                color='#13338E', lw=1.5, clip_on=False)
        ax.add_patch(arrow)
    if dead_end is not None:
        start = x_vals[-1] - dead_end
        de0 = np.clip(start, x0, x1)
        de1 = np.clip(start + dead_end, x0, x1)
        arrow = FancyArrowPatch((de0, y_arrow), (de1, y_arrow),
                                arrowstyle='<|-|>', mutation_scale=20,
                                color='#13338E', lw=1.5, clip_on=False)
        ax.add_patch(arrow)

    ax.legend(fontsize=10, loc='best')
    base = os.path.splitext(os.path.basename(file))[0]
    plt.savefig(os.path.join(output_folder, f"{base}_transferlength.pdf"), format='pdf')
    plt.savefig(os.path.join(output_folder, f"{base}_transferlength.png"), format='png')
    plt.close()
    return fig





class EnhancedTLCGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transfer Length Calculator")
        self.setGeometry(100, 100, 1200, 900)
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        self.file_path = None
        self.df = None           # <<<<<<<< NEU
        self.selected_time = None
        self.selected_row = None
        self.output_folder = os.path.join(os.getcwd(), "results")
        os.makedirs(self.output_folder, exist_ok=True)
        self.results = {}

        self.init_opening_screen()

    from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
    from PyQt5.QtCore import Qt
    from PyQt5.QtGui import QFont

    def init_opening_screen(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(30)

        # Titel
        title = QLabel("Transfer Length Calculator")
        title.setAlignment(Qt.AlignLeft)
        title.setStyleSheet("font-size: 48px; font-weight: 300;")
        layout.addWidget(title)

        # Beschreibungen (englisch / deutsch) – jetzt breiter
        desc_layout = QHBoxLayout()
        desc_layout.setSpacing(60)

        eng = QLabel(
            "With the TLC application, you can determine the anchorage lengths of CFRP strands in concrete<br>"
            "from DFOS strain-time data.<br><br>"
            "All details about the methodology and usage are described in the paper:<br>"
            "<i>Experimental Analysis of the Transfer Lengths of Prestressed CFRP Strands with Distributed Fiber Optic Sensors<br></i>"
            "María Serrano-Mesa et al. (2025)"
        )
        eng.setTextFormat(Qt.RichText)
        eng.setWordWrap(True)
        eng.setAlignment(Qt.AlignTop | Qt.AlignJustify)
        eng.setStyleSheet("font-size: 20px; font-weight: 300;")
        eng.setFixedWidth(520)
        eng.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        desc_layout.addWidget(eng, 1)

        deu = QLabel(
            "Mit der Anwendung TLC können Sie Verankerungslängen von CFK-Litzen in Beton<br>"
            "aus DFOS-Dehnungs-Zeit-Daten bestimmen.<br><br>"
            "Alle Informationen zur Methodik und Anwendung finden Sie im Paper:<br>"
            "<i>Experimental Analysis of the Transfer Lengths of Prestressed CFRP Strands with Distributed Fiber Optic Sensors<br></i>"
            "María Serrano-Mesa et al. (2025)"
        )
        deu.setTextFormat(Qt.RichText)
        deu.setWordWrap(True)
        deu.setAlignment(Qt.AlignTop | Qt.AlignJustify)
        deu.setStyleSheet("font-size: 20px; font-weight: 300;")
        deu.setFixedWidth(520)
        deu.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        desc_layout.addWidget(deu, 1)

        layout.addLayout(desc_layout)

        # Kompakte QTableWidget-Tabelle mit Daten
        table = QTableWidget()
        table.setRowCount(7)  # 1 Kopf + 6 Datenzeilen (inkl. Platzhalter)
        table.setColumnCount(6)

        headers = ["x₀", "x₁", "xᵢ", "⋯", "xₘ", "t"]
        table.setHorizontalHeaderLabels(headers)

        data = [
            ["ε₀,₀", "ε₀,₁", "ε₀,ⱼ", "⋯", "ε₀,ₘ", "t₀"],
            ["ε₁,₀", "ε₁,₁", "ε₁,ⱼ", "⋯", "ε₁,ₘ", "t₁"],
            ["εᵢ,₀", "εᵢ,₁", "εᵢ,ⱼ", "⋯", "εᵢ,ₘ", "tᵢ"],
            ["⋮", "⋮", "⋮", "⋱", "⋮", "⋮"],
            ["εₙ,₀", "εₙ,₁", "εₙ,ⱼ", "⋯", "εₙ,ₘ", "tₙ"]
        ]

        # Fülle Tabelle (erste Zeile Kopf ist schon gesetzt)
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, val in enumerate(row_data):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                font = QFont("Cambria Math")
                font.setPointSize(10)
                item.setFont(font)
                table.setItem(row_idx, col_idx, item)

        # Kopfzeilen formatieren
        header_font = QFont("Cambria Math")
        header_font.setBold(True)
        header_font.setPointSize(10)
        for col in range(table.columnCount()):
            header_item = table.horizontalHeaderItem(col)
            header_item.setFont(header_font)
            header_item.setTextAlignment(Qt.AlignCenter)

        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setFocusPolicy(Qt.NoFocus)

        # Spaltenbreite automatisch an Inhalt anpassen
        table.resizeColumnsToContents()

        # Gesamte Breite berechnen und fixieren (inkl. vertikaler Header-Breite)
        total_width = sum([table.columnWidth(i) for i in range(table.columnCount())]) + table.verticalHeader().width()
        table.setFixedWidth(total_width)

        # Zeilenhöhe kompakt setzen
        for row in range(table.rowCount()):
            table.setRowHeight(row, 18)

        # Tabelle horizontal zentrieren mit Spacer links und rechts
        table_container = QWidget()
        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(table)
        hbox.addStretch(1)
        hbox.setContentsMargins(0, 0, 0, 0)
        table_container.setLayout(hbox)

        layout.addWidget(table_container)

        # Bildunterschriften (englisch / deutsch nebeneinander)
        cap_layout = QHBoxLayout()
        cap_layout.setSpacing(60)

        cap_eng = QLabel(
            "<span style='font-size:18px;font-style:italic;'>"
            "To use the program, the Excel file must be formatted as shown:<br>"
            "Columns 1–(n-1): Strain values<br>Column n: Time</span>"
        )
        cap_eng.setTextFormat(Qt.RichText)
        cap_eng.setWordWrap(True)
        cap_eng.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        cap_eng.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        cap_layout.addWidget(cap_eng, 1)

        cap_deu = QLabel(
            "<span style='font-size:18px;font-style:italic;'>"
            "Um das Programm zu verwenden, muss die Excel-Datei wie abgebildet formatiert sein:<br>"
            "Spalten 1–(n-1): Dehnungswerte<br>Spalte n: Zeit</span>"
        )
        cap_deu.setTextFormat(Qt.RichText)
        cap_deu.setWordWrap(True)
        cap_deu.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        cap_deu.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        cap_layout.addWidget(cap_deu, 1)

        layout.addLayout(cap_layout)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(20)
        b = QPushButton("Select File")
        b.setFixedHeight(50)
        b.setMinimumWidth(220)
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(self.select_file)
        btn_layout.addWidget(b)
        layout.addLayout(btn_layout)

        # Footer
        footer = QLabel(
            "Ein Programm des Fachgebiets Entwerfen und Konstruieren – Massivbau, TU Berlin.\n"
            "Alle Rechte vorbehalten. Autoren: María Serrano-Mesa, Paul Merz, Oliver Disse"
        )
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("""
            font-size: 14px;
            font-weight: 300;
            color: rgba(19,51,142,0.7);
        """)
        layout.addWidget(footer)

        widget.setLayout(layout)
        self.stacked_widget.addWidget(widget)
        self.stacked_widget.setCurrentWidget(widget)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_name:
            self.file_path = file_name
            self.df = read_excel(file_name,self)  # <<<<<<<< Nur hier wird geladen!
            if self.df is None:
                QMessageBox.critical(self, "Error", "Could not read file.")
                return
            self.init_time_selection_screen()
        else:
            QMessageBox.warning(self, "Warning", "No file selected.")

    from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QSizePolicy

    from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QSizePolicy

    def init_time_selection_screen(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(30)

        # Überschrift (deutsch/englisch)
        headline_layout = QHBoxLayout()
        headline_layout.setSpacing(50)

        headline_eng = QLabel("""
            <h3 style="margin-bottom: 0.2em; text-align: left;">Time Selection for Analysis</h3>
        """)
        headline_eng.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        headline_eng.setStyleSheet("font-size: 24px; font-weight: 400; margin-bottom: 0.1em;")
        headline_eng.setMaximumWidth(500)
        headline_eng.setWordWrap(True)
        headline_eng.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        headline_layout.addWidget(headline_eng, 1)

        headline_deu = QLabel("""
            <h3 style="margin-bottom: 0.2em; text-align: left;">Zeitwahl für die Auswertung</h3>
        """)
        headline_deu.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        headline_deu.setStyleSheet("font-size: 24px; font-weight: 400; margin-bottom: 0.1em;")
        headline_deu.setMaximumWidth(500)
        headline_deu.setWordWrap(True)
        headline_deu.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        headline_layout.addWidget(headline_deu, 1)

        layout.addLayout(headline_layout)

        # Beschreibung englisch / deutsch nebeneinander mit Blocksatz und 1.5 Zeilenabstand
        desc_layout = QHBoxLayout()
        desc_layout.setSpacing(50)

        eng = QLabel()
        eng.setTextFormat(Qt.RichText)
        eng.setText("""
            <p style="text-align: justify; line-height: 1.5; word-break: break-word;">
                Please choose how you would like to determine the relevant analysis time:<br>
                • <b>Integral peak (just before crack):</b> Finds the moment when the integral of the strain distribution is maximal, which usually occurs just before the first major crack.<br>
                • <b>First measurement:</b> Uses the timestamp of the very first recorded data point.<br>
                • <b>Manual entry:</b> Enter a custom time (in seconds) for the analysis.
            </p>
        """)
        eng.setAlignment(Qt.AlignTop)
        eng.setStyleSheet("font-size: 18px; font-weight: 300;")
        eng.setMaximumWidth(500)
        eng.setWordWrap(True)
        eng.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        desc_layout.addWidget(eng, 1)

        deu = QLabel()
        deu.setTextFormat(Qt.RichText)
        deu.setText("""
            <p style="text-align: justify; line-height: 1.5; word-break: break-word;">
                Bitte wählen Sie, wie der relevante Zeitpunkt für die Auswertung bestimmt werden soll:<br>
                • <b>Integral-Spitze (unmittelbar vor Riss):</b> Sucht das Maximum des Integrals der Dehnungsverteilung – meist kurz vor dem ersten Hauptriss.<br>
                • <b>Erste Messung:</b> Verwendet den Zeitstempel des allerersten Messpunkts.<br>
                • <b>Manuelle Eingabe:</b> Geben Sie einen gewünschten Zeitpunkt (in Sekunden) selbst ein.
            </p>
        """)
        deu.setAlignment(Qt.AlignTop)
        deu.setStyleSheet("font-size: 18px; font-weight: 300;")
        deu.setMaximumWidth(500)
        deu.setWordWrap(True)
        deu.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        desc_layout.addWidget(deu, 1)

        layout.addLayout(desc_layout)

        # Buttons nur auf Englisch
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(20)
        button_defs = [
            ("Integral Peak", "Uses the time of maximum strain integral.", self.select_time_by_integral),
            ("First Measurement", "Uses first data timestamp.", self.select_time_first_row),
            ("Manual Entry", "Enter custom time in seconds.", self.manual_time_input)
        ]
        for text, tip, slot in button_defs:
            b = QPushButton(text)
            b.setFixedHeight(50)
            b.setCursor(Qt.PointingHandCursor)
            b.setToolTip(tip)
            b.clicked.connect(slot)
            btn_layout.addWidget(b)
        layout.addLayout(btn_layout)

        # Footer ohne Autoren
        footer = QLabel(
            "Ein Programm des Fachgebiets Entwerfen und Konstruieren – Massivbau, TU Berlin.\n"
            "Alle Rechte vorbehalten."
        )
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("""
            font-size: 14px;
            font-weight: 300;
            color: rgba(19,51,142,0.7);
        """)
        layout.addWidget(footer)

        widget.setLayout(layout)
        self.stacked_widget.addWidget(widget)
        self.stacked_widget.setCurrentWidget(widget)

    def select_time_by_integral(self):
        df = self.df   # <<<<<<<
        if df is None:
            QMessageBox.critical(self, "Error", "Data not loaded.")
            return
        cols = df.columns[:-1]
        df_clean = df.dropna(subset=cols, how='all')
        _, max_t = plot_integral_with_max(df, self.output_folder, "integral_plot.pdf")
        if max_t is None:
            QMessageBox.critical(self, "Error", "Integral plot failed.")
            return
        self.selected_time = max_t
        df_clean['time_diff'] = (df_clean[df.columns[-1]] - max_t).abs()
        self.selected_row = df_clean.loc[df_clean['time_diff'].idxmin()].copy()
        QMessageBox.information(self, "Time Determined", f"t = {self.selected_time}")
        self.show_integral_plot_screen()

    def show_integral_plot_screen(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        # Überschrift englisch/deutsch nebeneinander
        title_layout = QHBoxLayout()
        title_layout.setSpacing(50)

        lbl_eng = QLabel(f"Integral Plot (t = {self.selected_time:.3f} s)")
        lbl_eng.setAlignment(Qt.AlignCenter)
        lbl_eng.setStyleSheet("font-size: 24px; font-weight: 300;")
        title_layout.addWidget(lbl_eng, 1)

        lbl_deu = QLabel(f"Integral-Plot (t = {self.selected_time:.3f} s)")
        lbl_deu.setAlignment(Qt.AlignCenter)
        lbl_deu.setStyleSheet("font-size: 24px; font-weight: 300;")
        title_layout.addWidget(lbl_deu, 1)

        layout.addLayout(title_layout)

        # Beschreibung englisch / deutsch nebeneinander
        desc_layout = QHBoxLayout()
        desc_layout.setSpacing(50)

        eng = QLabel()
        eng.setTextFormat(Qt.RichText)
        eng.setText("""
            <p style="text-align: justify; line-height: 1.5; word-break: break-word;">
                This plot shows the integral of the strain distribution along the sensor length as a function of time.<br>
                The vertical line marks the time when the integral is maximal, usually just before the first crack forms.
            </p>
        """)
        eng.setAlignment(Qt.AlignTop)
        eng.setStyleSheet("font-size: 18px; font-weight: 300;")
        eng.setMaximumWidth(500)
        eng.setWordWrap(True)
        desc_layout.addWidget(eng, 1)

        deu = QLabel()
        deu.setTextFormat(Qt.RichText)
        deu.setText("""
            <p style="text-align: justify; line-height: 1.5; word-break: break-word;">
                Dieser Plot zeigt das Integral der Dehnungsverteilung über die Messlänge in Abhängigkeit von der Zeit.<br>
                Die vertikale Linie markiert den Zeitpunkt, an dem das Integral sein Maximum erreicht – in der Regel kurz vor dem ersten Riss.
            </p>
        """)
        deu.setAlignment(Qt.AlignTop)
        deu.setStyleSheet("font-size: 18px; font-weight: 300;")
        deu.setMaximumWidth(500)
        deu.setWordWrap(True)
        desc_layout.addWidget(deu, 1)

        layout.addLayout(desc_layout)

        # Plot-Bild
        img = os.path.join(self.output_folder, "integral_plot.png")
        if os.path.exists(img):
            pix = QPixmap(img)
            pl = QLabel()
            pl.setPixmap(pix)
            pl.setAlignment(Qt.AlignCenter)
            layout.addWidget(pl)
        else:
            not_found = QLabel("Plot not found.")
            not_found.setAlignment(Qt.AlignCenter)
            layout.addWidget(not_found)

        # Buttons nebeneinander
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(20)
        for text, slot in [
            ("Save Plot", self.save_integral_plot),
            ("Continue", self.init_analysis_dashboard)
        ]:
            b = QPushButton(text)
            b.setFixedHeight(50)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(slot)
            btn_layout.addWidget(b)
        layout.addLayout(btn_layout)

        widget.setLayout(layout)
        self.stacked_widget.addWidget(widget)
        self.stacked_widget.setCurrentWidget(widget)

    def save_integral_plot(self):
        base = os.path.join(self.output_folder, "integral_plot")
        src_png = base + ".png"
        src_pdf = base + ".pdf"
        if not (os.path.exists(src_png) or os.path.exists(src_pdf)):
            QMessageBox.warning(self, "Warning", "No plot to save.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Integral Plot", "integral_plot", "PNG Files (*.png);;PDF Files (*.pdf)"
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        src = src_png if ext == ".png" else src_pdf if ext == ".pdf" else None
        if src is None:
            QMessageBox.warning(self, "Warning", "Choose .png or .pdf")
            return
        try:
            shutil.copy(src, path)
            QMessageBox.information(self, "Saved", f"Saved as {ext}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Save failed: {e}")

    def select_time_first_row(self):
        df = self.df  # <<<<<<<
        if df is None:
            QMessageBox.critical(self, "Error", "Data not loaded.")
            return
        first = df.iloc[0]
        self.selected_time = first[df.columns[-1]]
        self.selected_row = first.copy()
        QMessageBox.information(self, "Time Selected", f"t = {self.selected_time}")
        self.init_analysis_dashboard()

    def manual_time_input(self):
        df = self.df  # <<<<<<<
        if df is None:
            QMessageBox.critical(self, "Error", "Data not loaded.")
            return
        s, ok = QInputDialog.getText(self, "Manual Time Input", "Enter time in seconds:")
        if ok:
            try:
                val = float(s)
                df['time_diff'] = (df[df.columns[-1]] - val).abs()
                self.selected_row = df.loc[df['time_diff'].idxmin()].copy()
                time_col = df.columns[-1]
                self.selected_time = self.selected_row[time_col]  # <-- Der echte Tabellen-Zeitwert!
                QMessageBox.information(self, "Time Selected", f"t = {self.selected_time}")
                self.init_analysis_dashboard()
            except ValueError:
                QMessageBox.warning(self, "Warning", "Invalid format.")

    # Dashboard mit EPS/L_OL-Auswahl, Zeit ist fix
    # ... innerhalb deiner EnhancedTLCGUI-Klasse ...

    def init_analysis_dashboard(self):
        class DashboardWidget(QWidget):
            def __init__(dash_self, parent_gui):
                super().__init__()
                dash_self.parent_gui = parent_gui
                dash_self.setContentsMargins(0, 0, 0, 0)
                dash_self.layout = QVBoxLayout()
                dash_self.layout.setContentsMargins(30, 30, 30, 30)
                dash_self.layout.setSpacing(12)

                hdr = QLabel(f"Transfer Length Analysis at t = {parent_gui.selected_time:.3f} s")
                hdr.setAlignment(Qt.AlignCenter)
                hdr.setStyleSheet("font-size: 22px; font-weight: 400; margin-bottom: 4px; margin-top:0;")
                dash_self.layout.addWidget(hdr)

                # --- TOP ROW: Eingabe links, Empfehlung rechts ---
                top_row = QHBoxLayout()
                # Eingabe links
                inputbox = QHBoxLayout()
                inputbox.addWidget(QLabel("Δε<sub>c</sub> [‰]:"), stretch=0)
                dash_self.eps_input = QLineEdit(str(getattr(parent_gui, "current_eps", 0.023)))
                dash_self.eps_input.setMaximumWidth(60)
                inputbox.addWidget(dash_self.eps_input, stretch=0)
                inputbox.addSpacing(10)
                inputbox.addWidget(QLabel("l<sub>ol</sub> [mm]:"), stretch=0)
                dash_self.lol_input = QLineEdit(str(getattr(parent_gui, "current_lol", 17)))
                dash_self.lol_input.setMaximumWidth(60)
                inputbox.addWidget(dash_self.lol_input, stretch=0)
                inputbox.addSpacing(16)

                dash_self.confirm_btn = QPushButton("Confirm")
                dash_self.confirm_btn.setCursor(Qt.PointingHandCursor)
                dash_self.confirm_btn.clicked.connect(dash_self.on_confirm)
                dash_self.confirm_btn.setMaximumWidth(200)
                inputbox.addWidget(dash_self.confirm_btn, stretch=0)
                inputbox.addStretch()
                top_row.addLayout(inputbox, stretch=2)

                # Empfehlung rechts (Schriftgröße angepasst)
                recommendation = QLabel(
                    "<span style='font-size:18px;'><b>Recommended:<br>"
                    "Δε<sub>c</sub>=<b>0.023</b>‰, l<sub>ol</sub>=<b>17</b>mm (embedded/einbetonierte Sensoren)<br>"
                    "Δε<sub>c</sub>=<b>0.020</b>‰, l<sub>ol</sub>=<b>16</b>mm (glued on surface/auf Oberfläche geklebte Sensoren) &nbsp; "
                    "<span style='font-size:14px;color:#888;'>(Serrano-Mesa et al. 2025)</span></span>"
                )

                recommendation.setTextFormat(Qt.RichText)
                recommendation.setStyleSheet("margin-left:36px; font-size: 18px;")  # wie Rest
                top_row.addWidget(recommendation, stretch=3)
                dash_self.layout.addLayout(top_row)

                # Fehlerlabel
                dash_self.error_label = QLabel("")
                dash_self.error_label.setStyleSheet("color:red; font-size:15px; margin-bottom:0;")
                dash_self.layout.addWidget(dash_self.error_label)

                # Plot ohne Erklärungstext
                plot_and_text_layout = QHBoxLayout()

                dash_self.analysis_plot_label = QLabel()
                dash_self.analysis_plot_label.setAlignment(Qt.AlignCenter)
                dash_self.analysis_plot_label.setStyleSheet(
                    "background: white; border:1px solid #ddd; min-height:300px;")
                plot_and_text_layout.addWidget(dash_self.analysis_plot_label, 1)

                dash_self.layout.addLayout(plot_and_text_layout, stretch=5)

                # Ergebnisse-Tabelle
                dash_self.analysis_table = QTableWidget(0, 5)
                dash_self.analysis_table.setHorizontalHeaderLabels(
                    ["Time [s]", "Δε₍c₎ [‰]", "l₍ol₎ [mm]", "Live End [mm]", "Dead End [mm]"]
                )
                dash_self.analysis_table.horizontalHeader().setStyleSheet("font-weight: 400; font-size: 18px;")
                dash_self.layout.addWidget(dash_self.analysis_table, stretch=1)

                # BUTTONS unten
                btns = QHBoxLayout()

                btn_save = QPushButton("Save All as Excel")
                btn_save.setCursor(Qt.PointingHandCursor)
                btn_save.clicked.connect(parent_gui.save_dashboard_results)
                btns.addWidget(btn_save)

                btn_save_plot = QPushButton("Save Plot")
                btn_save_plot.setCursor(Qt.PointingHandCursor)
                btn_save_plot.clicked.connect(dash_self.save_current_plot)
                btns.addWidget(btn_save_plot)

                btn_new_start = QPushButton("New Start")
                btn_new_start.setCursor(Qt.PointingHandCursor)
                btn_new_start.clicked.connect(dash_self.new_start)
                btns.addWidget(btn_new_start)

                btn_back_start = QPushButton("Start with new time")
                btn_back_start.setCursor(Qt.PointingHandCursor)
                # Neu: Direkt zur Zeitauswahl!
                btn_back_start.clicked.connect(parent_gui.init_time_selection_screen)
                btns.addWidget(btn_back_start)

                dash_self.layout.addLayout(btns)
                dash_self.setLayout(dash_self.layout)

                # Ergebnisliste initialisieren, falls nicht vorhanden
                if not hasattr(parent_gui, 'results_list'):
                    parent_gui.results_list = []

                dash_self.on_confirm()

            def resizeEvent(dash_self, event):
                dash_self.on_confirm()
                super().resizeEvent(event)

            def on_confirm(dash_self):
                try:
                    eps = float(dash_self.eps_input.text())
                    l_ol = float(dash_self.lol_input.text())
                except ValueError:
                    dash_self.error_label.setText("Both values must be valid numbers!")
                    return

                if eps <= 0 or l_ol <= 0:
                    dash_self.error_label.setText("Both values must be greater than zero!")
                    return
                else:
                    dash_self.error_label.setText("")

                parent_gui = dash_self.parent_gui
                parent_gui.current_eps = eps
                parent_gui.current_lol = l_ol

                # --- Robuste Auswahl der numerischen Positionsspalten ---
                row = parent_gui.selected_row
                deformation_cols = []
                positions = []
                for c in row.index[:-1]:
                    try:
                        f = float(str(c).strip())
                        deformation_cols.append(c)
                        positions.append(f)
                    except Exception:
                        continue

                # Debug-Ausgaben (im Terminal sichtbar)
                print("DEBUG deformation_cols:", deformation_cols)
                print("DEBUG positions:", positions)
                print("DEBUG row.shape:", row.shape)
                print("DEBUG row:", row[deformation_cols].values)

                if not positions:
                    dash_self.error_label.setText("Keine numerischen Positionsspalten erkannt!")
                    return

                vals = row[deformation_cols].values
                pts = [[v, positions[i]] for i, v in enumerate(vals) if not np.isnan(v)]
                y_arr = np.array([p[0] for p in pts])
                if y_arr.size == 0:
                    dash_self.error_label.setText("Keine Daten in der Zeile!")
                    return

                # --- Histogramm-Berechnung und Bin-Analyse ---
                bins = np.arange(0, np.nanmax(y_arr) + eps, eps)
                counts, _ = np.histogram(y_arr, bins=bins)
                digs = np.digitize(y_arr, bins) - 1
                bins_pts = [[] for _ in range(len(bins) - 1)]
                for i, pt in enumerate(pts):
                    bi = digs[i]
                    if 0 <= bi < len(bins_pts):
                        bins_pts[bi].append(pt)
                mb = np.argmax(counts)
                mb_vals = bins_pts[mb]
                max_edges = (bins[mb], bins[mb + 1])

                # --- Live End & Dead End Berechnung ---
                live_end, dead_end = None, None
                for i in range(len(mb_vals)):
                    valid = True
                    if i + 1 < len(mb_vals) and mb_vals[i][1] + l_ol <= mb_vals[i + 1][1]:
                        valid = False
                    if valid:
                        live_end = mb_vals[i][1]
                        break
                for i in range(len(mb_vals) - 1, -1, -1):
                    valid = True
                    if i - 1 >= 0 and mb_vals[i][1] >= mb_vals[i - 1][1] + l_ol:
                        valid = False
                    if valid:
                        dead_end = pts[-1][1] - mb_vals[i][1]
                        break

                parent_gui.results = {
                    "Time [s]": parent_gui.selected_time,
                    "Δε₍c₎ [‰]": eps,
                    "l₍ol₎ [mm]": l_ol,
                    "Live End [mm]": live_end,
                    "Dead End [mm]": dead_end
                }

                # --- Ergebnisliste pflegen ---
                already = False
                for i, rowx in enumerate(parent_gui.results_list):
                    if (
                            abs(rowx["Time [s]"] - parent_gui.selected_time) < 1e-6
                            and abs(rowx["Δε₍c₎ [‰]"] - eps) < 1e-9
                            and abs(rowx["l₍ol₎ [mm]"] - l_ol) < 1e-9
                    ):
                        parent_gui.results_list[i] = parent_gui.results.copy()
                        already = True
                        break
                if not already:
                    parent_gui.results_list.append(parent_gui.results.copy())

                dash_self.update_results_table()

                # --- Plot-Größe bestimmen ---
                aspect_ratio = 10 / 6
                widget_w = dash_self.analysis_plot_label.width()
                widget_h = dash_self.analysis_plot_label.height()
                if widget_w / aspect_ratio < widget_h:
                    width = max(6, widget_w / 100)
                    height = width / aspect_ratio
                else:
                    height = max(4, widget_h / 100)
                    width = height * aspect_ratio

                # --- Plot erzeugen und anzeigen ---
                plot_results(
                    pts, live_end, dead_end, l_ol, eps, max_edges,
                    parent_gui.output_folder, parent_gui.file_path, parent_gui.selected_time,
                    y_limits=None, figsize=(width, height)
                )
                base = os.path.splitext(os.path.basename(parent_gui.file_path))[0]
                img = os.path.join(parent_gui.output_folder, f"{base}_transferlength.png")
                if os.path.exists(img):
                    pix = QPixmap(img)
                    dash_self.analysis_plot_label.setPixmap(pix.scaled(
                        dash_self.analysis_plot_label.size(),
                        Qt.KeepAspectRatio, Qt.SmoothTransformation
                    ))
                else:
                    dash_self.analysis_plot_label.setText("Plot not found.")

            def update_results_table(dash_self):
                parent_gui = dash_self.parent_gui
                dash_self.analysis_table.setRowCount(len(parent_gui.results_list))
                for row_idx, result in enumerate(
                        parent_gui.results_list
                ):
                    for col, key in enumerate(
                            ["Time [s]", "Δε₍c₎ [‰]", "l₍ol₎ [mm]", "Live End [mm]", "Dead End [mm]"]
                    ):
                        item = QTableWidgetItem(str(result.get(key, "")))
                        item.setTextAlignment(Qt.AlignCenter)
                        dash_self.analysis_table.setItem(row_idx, col, item)
                # Spaltenbreite automatisch anpassen:
                dash_self.analysis_table.resizeColumnsToContents()

            def save_current_plot(dash_self):
                parent_gui = dash_self.parent_gui
                base = os.path.splitext(os.path.basename(parent_gui.file_path))[0]
                src = os.path.join(parent_gui.output_folder, f"{base}_transferlength.png")
                path, _ = QFileDialog.getSaveFileName(
                    dash_self, "Save Transfer Length Plot", "", "PNG Files (*.png)"
                )
                if path:
                    try:
                        shutil.copy(src, path)
                        QMessageBox.information(dash_self, "Saved", "Plot saved.")
                    except Exception as e:
                        QMessageBox.critical(dash_self, "Error", f"Save failed: {e}")

            def new_start(dash_self):
                # Leert die Tabelle und öffnet Dateiauswahl
                parent_gui = dash_self.parent_gui
                parent_gui.results_list = []
                parent_gui.file_path = None
                parent_gui.selected_time = None
                parent_gui.init_opening_screen()

        dash_widget = DashboardWidget(self)
        self.stacked_widget.addWidget(dash_widget)
        self.stacked_widget.setCurrentWidget(dash_widget)

    def save_dashboard_results(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel Results", "", "Excel Files (*.xlsx)"
        )
        if path:
            if hasattr(self, "results_list") and self.results_list:
                df = pd.DataFrame(self.results_list)
            else:
                df = pd.DataFrame([self.results])
            try:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "Saved", "Excel file saved.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Save failed: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setStyleSheet("""
        QWidget {
            background: white;
            color: #13338E;
            font-family: Georgia, serif;
            font-size: 20px;
            font-weight: 300;
        }
        QLabel {
            font-weight: 300;
        }
        QPushButton {
            background: #13338E;
            color: white;
            font-size: 22px;
            font-weight: 300;
            padding: 12px 24px;
            border: none;
            border-radius: 0px;  /* Eckig! */
        }
        QPushButton:hover {
            background: #0f296f;
        }
        QLineEdit {
            background: rgba(19,51,142,0.1);
            color: #13338E;
            font-size: 20px;
            font-weight: 300;
            padding: 8px;
            border: 1px solid rgba(19,51,142,0.5);
            border-radius: 4px;
        }
        QTableWidget {
            background: rgba(19,51,142,0.05);
            color: #13338E;
            font-size: 20px;
            font-weight: 300;
            gridline-color: rgba(19,51,142,0.3);
        }
    """)

    window = EnhancedTLCGUI()
    window.showMaximized()

    sys.exit(app.exec_())