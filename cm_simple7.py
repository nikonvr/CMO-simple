import sys
import numpy as np
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt
import mplcursors

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QTabWidget, QMessageBox,
    QScrollArea, QGroupBox, QSlider, QFileDialog, QStatusBar, QStyle
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QDoubleValidator, QIntValidator, QPalette, QColor, QIcon, QValidator

import pandas as pd
import os
import datetime

# Fonctions utilitaires robustes pour conversion numérique
def safe_str_to_float(text):
    """
    Convertit une chaîne en float de manière robuste, gérant point et virgule comme séparateur décimal.
    Retourne (valeur, succès)
    """
    if not text or not isinstance(text, str):
        return 0.0, False
    
    text = text.strip()
    if not text:
        return 0.0, False
    
    try:
        # Remplacer la virgule par un point pour la conversion
        normalized = text.replace(',', '.')
        # Supprimer les espaces
        normalized = normalized.replace(' ', '')
        # Gérer les cas avec plusieurs points/virgules (invalides)
        if normalized.count('.') > 1:
            return 0.0, False
        value = float(normalized)
        return value, True
    except (ValueError, AttributeError, TypeError):
        return 0.0, False

def safe_str_to_int(text):
    """
    Convertit une chaîne en int de manière robuste.
    Retourne (valeur, succès)
    """
    if not text or not isinstance(text, str):
        return 0, False
    
    text = text.strip()
    if not text:
        return 0, False
    
    try:
        # Pour int, on supprime d'abord les décimales si présentes
        normalized = text.replace(',', '.').replace(' ', '')
        # Si c'est un float, on le convertit en int
        if '.' in normalized:
            float_val = float(normalized)
            return int(round(float_val)), True
        value = int(normalized)
        return value, True
    except (ValueError, AttributeError, TypeError, OverflowError):
        return 0, False

def parse_empilement_string(emp_str):
    """
    Parse une chaîne d'empilement robuste, gérant point et virgule comme séparateur décimal.
    Format attendu: "1,0.5,1" ou "1,0,5,1" ou "1.5,2.3,1.0"
    Retourne (liste de floats, succès, message_erreur)
    """
    if not emp_str or not isinstance(emp_str, str):
        return [], True, ""
    
    emp_str = emp_str.strip()
    if not emp_str:
        return [], True, ""
    
    try:
        # Séparer par les virgules (séparateur de liste)
        parts = emp_str.split(',')
        emp_factors = []
        
        for i, part in enumerate(parts):
            part = part.strip()
            if not part:  # Ignorer les parties vides
                continue
            
            # Convertir chaque partie en float
            value, success = safe_str_to_float(part)
            if not success:
                return [], False, f"Valeur invalide à la position {i+1}: '{part}'"
            
            if value < 0:
                return [], False, f"Valeur négative non autorisée à la position {i+1}: {value}"
            
            emp_factors.append(value)
        
        return emp_factors, True, ""
    except Exception as e:
        return [], False, f"Erreur lors du parsing de l'empilement: {str(e)}"

# Style moderne et professionnel
GLOBAL_STYLESHEET = """
QMainWindow {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #f5f7fa, stop:1 #e8ecef);
}

QGroupBox {
    font-weight: bold;
    font-size: 11pt;
    color: #2c3e50;
    border: 2px solid #dee2e6;
    border-radius: 8px;
    margin-top: 1em;
    padding-top: 12px;
    background-color: white;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 8px;
    background-color: white;
    color: #495057;
}

QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #6c7ae0, stop:1 #5568d3);
    color: white;
    border: none;
    padding: 10px 20px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 10pt;
    min-height: 20px;
}

QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #7d8aef, stop:1 #6578e4);
    border: 1px solid #5a6fd8;
}

QPushButton:pressed {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #5568d3, stop:1 #4a5bc5);
}

QPushButton:disabled {
    background: #c0c0c0;
    color: #808080;
}

QPushButton#undoButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #5cb85c, stop:1 #4cae4c);
}

QPushButton#undoButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #6bc86b, stop:1 #5cb85c);
}

QPushButton#redoButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #5bc0de, stop:1 #46b8da);
}

QPushButton#redoButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #6bcfe8, stop:1 #5bc0de);
}

QPushButton#resetButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #f0ad4e, stop:1 #ec971f);
}

QPushButton#resetButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #f5b863, stop:1 #f0ad4e);
}

QLineEdit {
    border: 2px solid #ced4da;
    border-radius: 5px;
    padding: 6px 10px;
    background-color: white;
    font-size: 10pt;
    selection-background-color: #6c7ae0;
    selection-color: white;
}

QLineEdit:focus {
    border: 2px solid #6c7ae0;
    background-color: #f8f9ff;
}

QLineEdit[validity="valid"] {
    background-color: #d4edda;
    border: 2px solid #28a745;
}

QLineEdit[validity="invalid"] {
    background-color: #f8d7da;
    border: 2px solid #dc3545;
}

QLineEdit[validity="intermediate"] {
    background-color: #fff3cd;
    border: 2px solid #ffc107;
}

QScrollArea {
    border: none;
    background: transparent;
}

QScrollBar:vertical {
    border: none;
    background: #f1f3f5;
    width: 12px;
    border-radius: 6px;
}

QScrollBar::handle:vertical {
    background: #adb5bd;
    border-radius: 6px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background: #868e96;
}

QScrollBar:horizontal {
    border: none;
    background: #f1f3f5;
    height: 12px;
    border-radius: 6px;
}

QScrollBar::handle:horizontal {
    background: #adb5bd;
    border-radius: 6px;
    min-width: 20px;
}

QScrollBar::handle:horizontal:hover {
    background: #868e96;
}

QSlider::groove:horizontal {
    border: 1px solid #ced4da;
    height: 8px;
    background: #e9ecef;
    border-radius: 4px;
}

QSlider::handle:horizontal {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #6c7ae0, stop:1 #5568d3);
    border: 2px solid #495057;
    width: 20px;
    height: 20px;
    margin: -6px 0;
    border-radius: 10px;
}

QSlider::handle:horizontal:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #7d8aef, stop:1 #6578e4);
}

QCheckBox {
    font-size: 10pt;
    color: #495057;
    spacing: 8px;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border: 2px solid #ced4da;
    border-radius: 4px;
    background-color: white;
}

QCheckBox::indicator:hover {
    border: 2px solid #6c7ae0;
}

QCheckBox::indicator:checked {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #6c7ae0, stop:1 #5568d3);
    border: 2px solid #495057;
}

QTabWidget::pane {
    border: 2px solid #dee2e6;
    border-radius: 6px;
    background: white;
    top: -1px;
}

QTabBar::tab {
    background: #e9ecef;
    color: #495057;
    padding: 10px 20px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: 500;
}

QTabBar::tab:selected {
    background: white;
    color: #6c7ae0;
    border-bottom: 2px solid #6c7ae0;
}

QTabBar::tab:hover {
    background: #f8f9fa;
}

QStatusBar {
    background: #f8f9fa;
    color: #495057;
    border-top: 1px solid #dee2e6;
}
"""

# Configuration des champs d'entrée, utilisée pour la création et la réinitialisation
# Format: Label, Default, var_name, Validator, SliderConfig (min, max, display_mult), Tooltip
INPUT_CONFIGS = [
    ("Matériaux", [
        ("Indice du Superstrat (milieu incident):", "1.0", 'n_super', QDoubleValidator(0.1, 10.0, 4), (100, 400, 100), "Indice de réfraction réel du milieu d'où la lumière est incidente (superstrat)."),
        ("Matériau H (réel):", "2.25", 'nH_r', QDoubleValidator(0.1, 10.0, 4), (100, 400, 100), "Indice de réfraction réel du matériau haut indice (H)."),
        ("Matériau H (imaginaire):", "0.0001", 'nH_i', QDoubleValidator(0, 1.0, 5), None, "Partie imaginaire de l'indice de H (absorption). Mettre 0 pour non absorbant."),
        ("Matériau L (réel):", "1.48", 'nL_r', QDoubleValidator(0.1, 10.0, 4), (100, 400, 100), "Indice de réfraction réel du matériau bas indice (L)."),
        ("Matériau L (imaginaire):", "0.0001", 'nL_i', QDoubleValidator(0, 1.0, 5), None, "Partie imaginaire de l'indice de L (absorption). Mettre 0 pour non absorbant."),
        ("Substrat (indice réel):", "1.52", 'nSub_r', QDoubleValidator(0.1, 10.0, 4), (100, 400, 100), "Indice de réfraction réel du substrat."),
        ("Substrat (indice imaginaire):", "0.0", 'nSub_i', QDoubleValidator(0, 1.0, 5), None, "Partie imaginaire de l'indice du substrat (absorption). Mettre 0 pour non absorbant."),
    ]),
    ("Configuration de l'Empilement", [
        ("Longueur d'onde de centrage (nm):", "550", 'l0', QDoubleValidator(1, 2000, 1), (200, 1200, 1), "Longueur d'onde pour laquelle les épaisseurs QWOT sont calculées, sous l'incidence nominale."),
        ("Empilement (QWOT, ex: 1,0.5,1):", "1,1,1,1,1,2,1,1,1,1,1", 'emp_str', None, None, "Séquence des couches en multiples de QWOT (quart d'onde optique).\nExemple: '1,1' pour HL, '1,2,1' pour HLH (H=1xQWOT, L=2xQWOT, H=1xQWOT).\nLa première couche est H, puis L, etc."),
    ]),
    ("Paramètres Spectraux", [
        ("Intervalle spectral début (nm):", "400", 'l_range_deb', QDoubleValidator(1, 2000, 1), (200, 1000, 1),"Début de l'intervalle pour le tracé spectral."),
        ("Intervalle spectral fin (nm):", "700", 'l_range_fin', QDoubleValidator(1, 2000, 1), (400, 1500, 1), "Fin de l'intervalle pour le tracé spectral."),
        ("Pas spectral (nm):", "1", 'l_step', QDoubleValidator(0.01, 100, 2), (1, 1000, 100), "Pas de calcul pour le tracé spectral (ex: 0.1 pour 0.01nm)."),
    ]),
    ("Paramètres Angulaires", [
        ("Incidence nominale (design & tracé spectral) (°):", "0", 'inc', QDoubleValidator(0, 89.99, 2), (0, 8999, 100), "Angle d'incidence (dans le superstrat) pour le design QWOT et pour le tracé spectral fixe."),
        ("Intervalle angulaire début (°):", "0", 'a_range_deb', QDoubleValidator(0, 89.99, 2), (0, 8999, 100), "Début de l'intervalle pour le tracé angulaire."),
        ("Intervalle angulaire fin (°):", "89", 'a_range_fin', QDoubleValidator(0, 89.99, 2), (0, 8999, 100), "Fin de l'intervalle pour le tracé angulaire."),
        ("Pas angulaire (°):", "1", 'a_step', QDoubleValidator(0.01, 90, 2), (1, 500, 100), "Pas de calcul pour le tracé angulaire (ex: 0.1 pour 0.01°)."),
    ]),
    ("Options de Calcul et d'Affichage", []) # Placeholder for checkboxes
]

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Calcul d'empilement de couches minces (PyQt6) - V7 (Superstrat)")
        self.setGeometry(50, 50, 1400, 900)
        self.setStyleSheet(GLOBAL_STYLESHEET)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Prêt.")

        self.input_panel = QWidget()
        self.input_panel_layout = QVBoxLayout(self.input_panel)
        self.input_panel.setFixedWidth(550) # Slightly wider for new field

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_area.setWidget(self.scroll_content)
        self.input_form_layout = QVBoxLayout(self.scroll_content)

        self.entry_vars_qt = {}
        self.sliders_qt = {}
        self.slider_value_labels = {}
        self.default_input_values = {}
        
        # Système UNDO/REDO
        self.undo_history = []  # Historique des états (max 5)
        self.redo_history = []  # Pile pour REDO
        self.max_undo_steps = 5
        self.is_undoing = False  # Flag pour éviter de sauvegarder lors d'un UNDO
        self.is_redoing = False  # Flag pour éviter de sauvegarder lors d'un REDO

        self._setup_input_fields()
        
        # Barre d'outils avec UNDO/REDO
        undo_redo_layout = QHBoxLayout()
        self.undo_button = QPushButton("↶ UNDO")
        self.undo_button.setObjectName("undoButton")
        self.undo_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowBack))
        self.undo_button.setToolTip("Annuler la dernière modification (max 5)")
        self.undo_button.clicked.connect(self._undo)
        self.undo_button.setEnabled(False)
        undo_redo_layout.addWidget(self.undo_button)
        
        self.redo_button = QPushButton("↷ REDO")
        self.redo_button.setObjectName("redoButton")
        self.redo_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowForward))
        self.redo_button.setToolTip("Refaire la dernière annulation")
        self.redo_button.clicked.connect(self._redo)
        self.redo_button.setEnabled(False)
        undo_redo_layout.addWidget(self.redo_button)
        
        undo_redo_layout.addStretch()
        
        self.reset_button = QPushButton("🔄 Réinitialiser")
        self.reset_button.setObjectName("resetButton")
        self.reset_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogResetButton))
        self.reset_button.clicked.connect(self._reset_parameters_to_default)
        undo_redo_layout.addWidget(self.reset_button)
        
        self.input_form_layout.addLayout(undo_redo_layout)
        
        self.input_panel_layout.addWidget(self.scroll_area)
        self.main_layout.addWidget(self.input_panel)

        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)
        self._setup_plot_tabs()

        self.cursors_spectral = []
        self.cursors_angular = []
        
        self.recalculation_timer = QTimer(self)
        self.recalculation_timer.setSingleShot(True)
        self.recalculation_timer.timeout.connect(self._perform_recalculation_and_plot)
        
        # Timer pour sauvegarder l'état avec délai (évite trop de sauvegardes)
        self.save_state_timer = QTimer(self)
        self.save_state_timer.setSingleShot(True)
        self.save_state_timer.timeout.connect(self._save_state_to_history)
        
        self._reset_parameters_to_default(initial_load=True)
        
        # Sauvegarder l'état initial
        self._save_state_to_history()
    
    def _save_current_state(self):
        """Sauvegarde l'état actuel de tous les champs dans un dictionnaire."""
        state = {}
        # Sauvegarder tous les champs de saisie
        for var_name, entry in self.entry_vars_qt.items():
            state[var_name] = entry.text()
        # Sauvegarder les checkboxes
        state['plot_rs'] = self.plot_rs_checkbox.isChecked()
        state['plot_rp'] = self.plot_rp_checkbox.isChecked()
        state['plot_ts'] = self.plot_ts_checkbox.isChecked()
        state['plot_tp'] = self.plot_tp_checkbox.isChecked()
        state['autoscale_y'] = self.autoscale_y_checkbox.isChecked()
        state['substrat_fini'] = self.substrat_fini_checkbox.isChecked()
        state['export_excel'] = self.export_excel_checkbox.isChecked()
        return state
    
    def _restore_state(self, state):
        """Restaure un état sauvegardé."""
        self.is_undoing = True  # Empêcher la sauvegarde lors de la restauration
        # Restaurer tous les champs de saisie
        for var_name, value in state.items():
            if var_name in self.entry_vars_qt:
                entry = self.entry_vars_qt[var_name]
                entry.blockSignals(True)
                entry.setText(value)
                self._validate_line_edit_style(entry)
                entry.blockSignals(False)
                
                # Mettre à jour les sliders associés si nécessaire
                if var_name in self.sliders_qt:
                    slider = self.sliders_qt[var_name]
                    slider_cfg_found = None
                    for group_cfg in INPUT_CONFIGS:
                        for field_cfg in group_cfg[1]:
                            if field_cfg[2] == var_name:
                                slider_cfg_found = field_cfg[4]
                                break
                        if slider_cfg_found:
                            break
                    if slider_cfg_found:
                        display_multiplier = slider_cfg_found[2]
                        slider_value_label = self.slider_value_labels.get(var_name)
                        self._update_slider_from_lineedit(entry, slider, display_multiplier, var_name, slider_value_label, suppress_recalc=True)
        
        # Restaurer les checkboxes
        if 'plot_rs' in state:
            self.plot_rs_checkbox.blockSignals(True)
            self.plot_rs_checkbox.setChecked(state['plot_rs'])
            self.plot_rs_checkbox.blockSignals(False)
        if 'plot_rp' in state:
            self.plot_rp_checkbox.blockSignals(True)
            self.plot_rp_checkbox.setChecked(state['plot_rp'])
            self.plot_rp_checkbox.blockSignals(False)
        if 'plot_ts' in state:
            self.plot_ts_checkbox.blockSignals(True)
            self.plot_ts_checkbox.setChecked(state['plot_ts'])
            self.plot_ts_checkbox.blockSignals(False)
        if 'plot_tp' in state:
            self.plot_tp_checkbox.blockSignals(True)
            self.plot_tp_checkbox.setChecked(state['plot_tp'])
            self.plot_tp_checkbox.blockSignals(False)
        if 'autoscale_y' in state:
            self.autoscale_y_checkbox.blockSignals(True)
            self.autoscale_y_checkbox.setChecked(state['autoscale_y'])
            self.autoscale_y_checkbox.blockSignals(False)
        if 'substrat_fini' in state:
            self.substrat_fini_checkbox.blockSignals(True)
            self.substrat_fini_checkbox.setChecked(state['substrat_fini'])
            self.substrat_fini_checkbox.blockSignals(False)
        if 'export_excel' in state:
            self.export_excel_checkbox.blockSignals(True)
            self.export_excel_checkbox.setChecked(state['export_excel'])
            self.export_excel_checkbox.blockSignals(False)
        
        self.is_undoing = False
        # Déclencher le recalcul après restauration
        self._schedule_recalculation()
    
    def _save_state_to_history_delayed(self):
        """Déclenche une sauvegarde avec délai pour éviter trop de sauvegardes."""
        self.save_state_timer.stop()
        self.save_state_timer.start(300)  # 300ms de délai
    
    def _save_state_to_history(self):
        """Sauvegarde l'état actuel dans l'historique UNDO (si pas en train d'UNDO/REDO)."""
        if self.is_undoing or self.is_redoing:
            return
        
        current_state = self._save_current_state()
        
        # Ne pas sauvegarder si c'est le même état que le dernier
        if self.undo_history and current_state == self.undo_history[-1]:
            return
        
        # Ajouter à l'historique
        self.undo_history.append(current_state)
        
        # Limiter à max_undo_steps états
        if len(self.undo_history) > self.max_undo_steps:
            self.undo_history.pop(0)
        
        # Vider le redo quand on fait une nouvelle action
        self.redo_history.clear()
        
        # Mettre à jour les boutons
        self.undo_button.setEnabled(len(self.undo_history) > 1)
        self.redo_button.setEnabled(False)
    
    def _undo(self):
        """Annule la dernière action."""
        if len(self.undo_history) <= 1:
            return
        
        # Sauvegarder l'état actuel dans redo
        current_state = self._save_current_state()
        self.redo_history.append(current_state)
        
        # Retirer le dernier état de l'historique
        self.undo_history.pop()
        
        # Restaurer l'état précédent
        if self.undo_history:
            self.is_undoing = True
            previous_state = self.undo_history[-1].copy()
            self._restore_state(previous_state)
            self.is_undoing = False
        
        # Mettre à jour les boutons
        self.undo_button.setEnabled(len(self.undo_history) > 1)
        self.redo_button.setEnabled(len(self.redo_history) > 0)
        self.status_bar.showMessage("Action annulée", 2000)
    
    def _redo(self):
        """Refait la dernière action annulée."""
        if not self.redo_history:
            return
        
        # Sauvegarder l'état actuel dans l'historique undo
        current_state = self._save_current_state()
        if not self.undo_history or current_state != self.undo_history[-1]:
            self.undo_history.append(current_state)
            if len(self.undo_history) > self.max_undo_steps:
                self.undo_history.pop(0)
        
        # Restaurer l'état depuis redo
        state_to_restore = self.redo_history.pop()
        self.is_redoing = True
        self._restore_state(state_to_restore)
        self.is_redoing = False
        
        # Ajouter l'état restauré à l'historique undo
        if not self.undo_history or state_to_restore != self.undo_history[-1]:
            self.undo_history.append(state_to_restore.copy())
            if len(self.undo_history) > self.max_undo_steps:
                self.undo_history.pop(0)
        
        # Mettre à jour les boutons
        self.undo_button.setEnabled(len(self.undo_history) > 1)
        self.redo_button.setEnabled(len(self.redo_history) > 0)
        self.status_bar.showMessage("Action refaite", 2000)

    def _setup_input_fields(self):
        for group_title, fields in INPUT_CONFIGS:
            group_box = QGroupBox(group_title)
            group_layout = QGridLayout(group_box)
            current_row_in_group = 0

            if not fields and group_title == "Options de Calcul et d'Affichage": # Special handling for checkboxes group
                 self._setup_checkboxes(group_layout) # Pass layout to fill
            else:
                for label_text, default_value, var_name, validator, slider_config, tooltip_text in fields:
                    self.default_input_values[var_name] = default_value

                    label = QLabel(label_text)
                    label.setToolTip(tooltip_text)
                    entry = QLineEdit(default_value)
                    entry.setToolTip(tooltip_text)
                    if validator:
                        entry.setValidator(validator)
                    
                    group_layout.addWidget(label, current_row_in_group, 0)
                    
                    if slider_config:
                        slider_min, slider_max, display_multiplier = slider_config
                        slider = QSlider(Qt.Orientation.Horizontal)
                        slider.setRange(slider_min, slider_max)
                        slider.setToolTip(tooltip_text)
                        
                        slider_value_label = QLabel("")
                        self.slider_value_labels[var_name] = slider_value_label
                        
                        # Conversion robuste de la valeur par défaut
                        initial_float_val, success = safe_str_to_float(default_value)
                        if success:
                            slider.setValue(int(initial_float_val * display_multiplier))
                            if display_multiplier == 1:
                                slider_value_label.setText(f" ({int(initial_float_val)})")
                            else:
                                decimals = len(str(display_multiplier)) - 1
                                slider_value_label.setText(f" ({initial_float_val:.{decimals}f})")
                        else:
                            # Default to min if conversion fails
                            slider.setValue(slider_min)
                            default_display_val = slider_min / display_multiplier
                            if display_multiplier == 1:
                                slider_value_label.setText(f" ({int(default_display_val)})")
                            else:
                                decimals = len(str(display_multiplier)) - 1
                                slider_value_label.setText(f" ({default_display_val:.{decimals}f})")


                        self.sliders_qt[var_name] = slider
                        
                        slider.valueChanged.connect(lambda value, le=entry, mult=display_multiplier, vn=var_name, lbl=slider_value_label: 
                                                     self._update_lineedit_from_slider(value, le, mult, vn, lbl))
                        entry.editingFinished.connect(lambda le=entry, sl=slider, mult=display_multiplier, vn=var_name, lbl=slider_value_label: 
                                                      self._update_slider_from_lineedit(le, sl, mult, vn, lbl))
                        entry.textChanged.connect(lambda text, le=entry: self._validate_line_edit_style(le))
                        # Sauvegarder l'état avant modification
                        entry.textChanged.connect(lambda: self._save_state_to_history_delayed())


                        input_widget_layout = QHBoxLayout()
                        input_widget_layout.addWidget(entry, 2) 
                        input_widget_layout.addWidget(slider, 3) 
                        input_widget_layout.addWidget(slider_value_label, 1)
                        container_widget = QWidget()
                        container_widget.setLayout(input_widget_layout)
                        group_layout.addWidget(container_widget, current_row_in_group, 1)

                    else: 
                        entry.editingFinished.connect(self._schedule_recalculation)
                        entry.textChanged.connect(lambda text, le=entry: self._validate_line_edit_style(le))
                        # Sauvegarder l'état avant modification
                        entry.textChanged.connect(lambda: self._save_state_to_history_delayed())
                        group_layout.addWidget(entry, current_row_in_group, 1)

                    if var_name == 'emp_str':
                        entry.textChanged.connect(self.update_layers_count_qt) 
                    
                    self.entry_vars_qt[var_name] = entry
                    current_row_in_group += 1
            self.input_form_layout.addWidget(group_box)
        
        self.layers_count_label_qt = QLabel("Nombre de couches : 0")
        self.input_form_layout.addWidget(self.layers_count_label_qt)
        self.update_layers_count_qt()

    def _setup_checkboxes(self, layout): # layout is QGridLayout of the "Options" group
        self.plot_rs_checkbox = QCheckBox("Afficher Rs")
        self.plot_rs_checkbox.setChecked(True)
        self.plot_rs_checkbox.stateChanged.connect(self._schedule_recalculation)
        self.plot_rs_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.plot_rs_checkbox, 0, 0)

        self.plot_rp_checkbox = QCheckBox("Afficher Rp")
        self.plot_rp_checkbox.setChecked(True)
        self.plot_rp_checkbox.stateChanged.connect(self._schedule_recalculation)
        self.plot_rp_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.plot_rp_checkbox, 0, 1)

        self.plot_ts_checkbox = QCheckBox("Afficher Ts")
        self.plot_ts_checkbox.setChecked(True)
        self.plot_ts_checkbox.stateChanged.connect(self._schedule_recalculation)
        self.plot_ts_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.plot_ts_checkbox, 1, 0)

        self.plot_tp_checkbox = QCheckBox("Afficher Tp")
        self.plot_tp_checkbox.setChecked(True)
        self.plot_tp_checkbox.stateChanged.connect(self._schedule_recalculation)
        self.plot_tp_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.plot_tp_checkbox, 1, 1)
        
        self.autoscale_y_checkbox = QCheckBox("Échelle Y Automatique (Graph. Spectral/Angulaire)")
        self.autoscale_y_checkbox.setChecked(False) 
        self.autoscale_y_checkbox.stateChanged.connect(self._schedule_recalculation)
        self.autoscale_y_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.autoscale_y_checkbox, 2, 0, 1, 2)

        self.substrat_fini_checkbox = QCheckBox("Substrat fini (réflexions multiples face arrière)")
        self.substrat_fini_checkbox.stateChanged.connect(self._schedule_recalculation)
        self.substrat_fini_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.substrat_fini_checkbox, 3, 0, 1, 2)

        self.export_excel_checkbox = QCheckBox("Exporter vers Excel lors du calcul")
        self.export_excel_checkbox.stateChanged.connect(lambda: self._save_state_to_history_delayed())
        layout.addWidget(self.export_excel_checkbox, 4, 0, 1, 2)


    def _validate_line_edit_style(self, line_edit):
        validator = line_edit.validator()
        if validator:
            state, _, _ = validator.validate(line_edit.text(), 0)
            if state == QValidator.State.Acceptable:
                line_edit.setProperty("validity", "valid")
            elif state == QValidator.State.Intermediate:
                line_edit.setProperty("validity", "intermediate")
            else: 
                line_edit.setProperty("validity", "invalid")
        else: 
            line_edit.setProperty("validity", "")
        line_edit.style().unpolish(line_edit)
        line_edit.style().polish(line_edit)


    def _setup_plot_tabs(self):
        self.tab_spectral = QWidget()
        self.tabs.addTab(self.tab_spectral, "Graphique Spectral")
        tab_spectral_outer_layout = QVBoxLayout(self.tab_spectral)
        self.fig_spectral = Figure(figsize=(7, 5)) 
        self.canvas_spectral = FigureCanvasQTAgg(self.fig_spectral)
        tab_spectral_outer_layout.addWidget(NavigationToolbar(self.canvas_spectral, self))
        tab_spectral_outer_layout.addWidget(self.canvas_spectral)
        self.export_spectral_button = QPushButton("Exporter Graphique Spectral")
        self.export_spectral_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        self.export_spectral_button.clicked.connect(lambda: self._export_figure(self.fig_spectral, "graphique_spectral"))
        tab_spectral_outer_layout.addWidget(self.export_spectral_button)
        self.ax_spectral = self.fig_spectral.add_subplot(111)
        self.fig_spectral.tight_layout()

        self.tab_angular = QWidget()
        self.tabs.addTab(self.tab_angular, "Graphique Angulaire")
        tab_angular_outer_layout = QVBoxLayout(self.tab_angular)
        self.fig_angular = Figure(figsize=(7, 5))
        self.canvas_angular = FigureCanvasQTAgg(self.fig_angular)
        tab_angular_outer_layout.addWidget(NavigationToolbar(self.canvas_angular, self))
        tab_angular_outer_layout.addWidget(self.canvas_angular)
        self.export_angular_button = QPushButton("Exporter Graphique Angulaire")
        self.export_angular_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        self.export_angular_button.clicked.connect(lambda: self._export_figure(self.fig_angular, "graphique_angulaire"))
        tab_angular_outer_layout.addWidget(self.export_angular_button)
        self.ax_angular = self.fig_angular.add_subplot(111)
        self.fig_angular.tight_layout()

        self.tab_stack_vis = QWidget()
        self.tabs.addTab(self.tab_stack_vis, "Visualisation de l'Empilement")
        tab_stack_outer_layout = QVBoxLayout(self.tab_stack_vis)
        self.fig_stack_vis = Figure(figsize=(7, 5)) 
        self.canvas_stack_vis = FigureCanvasQTAgg(self.fig_stack_vis)
        tab_stack_outer_layout.addWidget(NavigationToolbar(self.canvas_stack_vis, self))
        tab_stack_outer_layout.addWidget(self.canvas_stack_vis)
        self.export_stack_button = QPushButton("Exporter Visualisation Empilement")
        self.export_stack_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        self.export_stack_button.clicked.connect(lambda: self._export_figure(self.fig_stack_vis, "visualisation_empilement"))
        tab_stack_outer_layout.addWidget(self.export_stack_button)
        self.ax_refractive_index_profile = self.fig_stack_vis.add_subplot(111)
        self.fig_stack_vis.tight_layout()

    def _export_figure(self, figure, default_filename_prefix):
        filename, _ = QFileDialog.getSaveFileName(
            self, 
            "Exporter le graphique", 
            f"{default_filename_prefix}.png", 
            "Images PNG (*.png);;Images SVG (*.svg);;Tous les fichiers (*)"
        )
        if filename:
            try:
                figure.savefig(filename, bbox_inches='tight')
                self.status_bar.showMessage(f"Graphique exporté vers : {filename}", 5000)
            except Exception as e:
                QMessageBox.critical(self, "Erreur d'Exportation", f"Impossible d'exporter le graphique : {e}")
                self.status_bar.showMessage("Erreur d'exportation du graphique.", 5000)


    def _reset_parameters_to_default(self, initial_load=False):
        self.status_bar.showMessage("Réinitialisation des paramètres...", 2000)
        for var_name, default_value in self.default_input_values.items():
            if var_name in self.entry_vars_qt:
                line_edit = self.entry_vars_qt[var_name]
                line_edit.blockSignals(True)
                line_edit.setText(default_value)
                self._validate_line_edit_style(line_edit)
                line_edit.blockSignals(False)

                if var_name in self.sliders_qt:
                    slider = self.sliders_qt[var_name]
                    # Trouver le display_multiplier pour ce var_name à partir de INPUT_CONFIGS
                    slider_cfg_found = None
                    for group_cfg in INPUT_CONFIGS:
                        for field_cfg in group_cfg[1]:
                            if field_cfg[2] == var_name:
                                slider_cfg_found = field_cfg[4] # slider_config
                                break
                        if slider_cfg_found:
                            break
                    
                    if slider_cfg_found:
                        display_multiplier = slider_cfg_found[2]
                        slider_value_label = self.slider_value_labels.get(var_name)
                        self._update_slider_from_lineedit(line_edit, slider, display_multiplier, var_name, slider_value_label, suppress_recalc=True)
        
        self.plot_rs_checkbox.setChecked(True)
        self.plot_rp_checkbox.setChecked(True)
        self.plot_ts_checkbox.setChecked(True)
        self.plot_tp_checkbox.setChecked(True)
        self.autoscale_y_checkbox.setChecked(False)
        self.substrat_fini_checkbox.setChecked(False)
        self.export_excel_checkbox.setChecked(False)

        if not initial_load:
             # Sauvegarder l'état avant réinitialisation pour UNDO
             self._save_state_to_history()
             self._schedule_recalculation()
        else:
             self._perform_recalculation_and_plot() 
        self.status_bar.showMessage("Paramètres réinitialisés.", 3000)


    def _update_lineedit_from_slider(self, value, line_edit, multiplier, var_name, slider_value_label):
        """Met à jour le line_edit depuis le slider avec gestion robuste."""
        try:
            # Sauvegarder l'état avant modification si le line_edit n'est pas bloqué
            if not line_edit.signalsBlocked():
                self._save_state_to_history_delayed()
            
            line_edit.blockSignals(True)
            
            # Calcul sécurisé de la valeur
            if multiplier == 0:
                multiplier = 1.0  # Éviter division par zéro
            actual_value = float(value) / float(multiplier)
            
            # Calcul du nombre de décimales
            decimals = 0
            if multiplier > 1:
                try:
                    decimals = int(np.log10(multiplier)) if multiplier > 0 else 2
                except (ValueError, OverflowError):
                    decimals = 2
            
            # Formatage
            if multiplier == 1 or decimals == 0:
                formatted_value = str(int(round(actual_value)))
                slider_label_text = f" ({int(round(actual_value))})"
            else:
                formatted_value = f"{actual_value:.{decimals}f}"
                slider_label_text = f" ({actual_value:.{decimals}f})"
            
            line_edit.setText(formatted_value)
            if slider_value_label:
                try:
                    slider_value_label.setText(slider_label_text)
                except:
                    pass  # Ignorer si l'affichage échoue
            
            self._validate_line_edit_style(line_edit)
            line_edit.blockSignals(False)
            self._schedule_recalculation()
        except Exception as e:
            # En cas d'erreur, restaurer l'état du line_edit
            try:
                line_edit.blockSignals(False)
            except:
                pass

    def _update_slider_from_lineedit(self, line_edit, slider, multiplier, var_name, slider_value_label, suppress_recalc=False):
        """Met à jour le slider depuis le line_edit avec gestion robuste."""
        try:
            value, success = safe_str_to_float(line_edit.text())
            if not success:
                return  # Ne rien faire si la conversion échoue
            
            # Calcul sécurisé
            if multiplier == 0:
                multiplier = 1.0  # Éviter division par zéro
            slider_value = int(round(float(value) * float(multiplier)))
            
            slider.blockSignals(True)
            slider.setValue(slider_value)
            slider.blockSignals(False)
            
            # Calcul sécurisé de la valeur affichée
            actual_value_from_slider = float(slider.value()) / float(multiplier)
            decimals = 0
            if multiplier > 1:
                try:
                    decimals = int(np.log10(multiplier)) if multiplier > 0 else 2
                except (ValueError, OverflowError):
                    decimals = 2

            if multiplier == 1 or decimals == 0:
                slider_label_text = f" ({int(round(actual_value_from_slider))})"
            else:
                slider_label_text = f" ({actual_value_from_slider:.{decimals}f})"
            
            if slider_value_label:
                try:
                    slider_value_label.setText(slider_label_text)
                except:
                    pass  # Ignorer si l'affichage échoue

        except (ValueError, TypeError, ZeroDivisionError, OverflowError):
            pass  # Ignorer les erreurs de conversion silencieusement
        except Exception:
            pass  # Ignorer toute autre erreur pour éviter les plantages 
        if not suppress_recalc:
            self._schedule_recalculation() 

    def _schedule_recalculation(self, initial=False):
        """Planifie un recalcul avec optimisation (évite les recalculs multiples)."""
        try:
            if initial: 
                self._perform_recalculation_and_plot()
            else:
                # Arrêter le timer existant pour éviter les recalculs multiples
                if self.recalculation_timer.isActive():
                    self.recalculation_timer.stop()
                self.status_bar.showMessage("Préparation du calcul...", 1000)
                # Timer plus court pour plus de réactivité (150ms au lieu de 200ms)
                self.recalculation_timer.start(150)
        except Exception:
            # Si le scheduling échoue, essayer quand même le calcul immédiat
            try:
                self._perform_recalculation_and_plot()
            except:
                pass  # Dernier recours - ignorer 

    def update_layers_count_qt(self):
        try:
            emp_entry = self.entry_vars_qt.get('emp_str')
            if not emp_entry:
                self.layers_count_label_qt.setText("Nombre de couches : 0")
                return
            
            emp_str = emp_entry.text()
            if not emp_str or not emp_str.strip():
                num_layers = 0
            else:
                # Utiliser le parsing robuste pour compter correctement
                emp_factors, success, _ = parse_empilement_string(emp_str)
                if success:
                    num_layers = len(emp_factors)
                else:
                    # Si le parsing échoue, essayer de compter les virgules + 1
                    num_layers = max(0, emp_str.count(',') + (1 if emp_str.strip() else 0))
            
            self.layers_count_label_qt.setText(f"Nombre de couches : {num_layers}")
        except Exception as e:
            # En cas d'erreur, afficher un message neutre
            try:
                self.layers_count_label_qt.setText("Nombre de couches : ?")
            except:
                pass  # Ignorer si même ça échoue

    def clear_plot_cursors(self, main_cursor_list):
        for group_of_cursors in main_cursor_list: 
            for c_item in group_of_cursors:       
                try:
                    c_item.remove()
                except Exception:
                    pass
        main_cursor_list.clear()
            
    def plot_spectral_data(self, res, inc_val, n_super_val):
        self.ax_spectral.clear()
        self.clear_plot_cursors(self.cursors_spectral) 
        
        spectral_lines = []
        if res['l'].size > 0:
            if self.plot_rs_checkbox.isChecked() and res['Rs_s'].size == res['l'].size: 
                l1, = self.ax_spectral.plot(res['l'], res['Rs_s'], label='Rs', linestyle='-')
                spectral_lines.append(l1)
            if self.plot_rp_checkbox.isChecked() and res['Rp_s'].size == res['l'].size:
                l2, = self.ax_spectral.plot(res['l'], res['Rp_s'], label='Rp', linestyle='--')
                spectral_lines.append(l2)
            if self.plot_ts_checkbox.isChecked() and res['Ts_s'].size == res['l'].size:
                l3, = self.ax_spectral.plot(res['l'], res['Ts_s'], label='Ts', linestyle='-')
                spectral_lines.append(l3)
            if self.plot_tp_checkbox.isChecked() and res['Tp_s'].size == res['l'].size:
                l4, = self.ax_spectral.plot(res['l'], res['Tp_s'], label='Tp', linestyle='--')
                spectral_lines.append(l4)

            self.ax_spectral.set_xlabel('Longueur d\'onde (nm)')
            self.ax_spectral.set_ylabel('Reflectance / Transmittance')
            self.ax_spectral.set_title(f"Tracé spectral (n_super={n_super_val:.2f}, incidence {inc_val:.1f}°)")
            self.ax_spectral.grid(True, which='major', color='grey', linestyle='-', linewidth=0.7)
            self.ax_spectral.grid(True, which='minor', color='lightgrey', linestyle=':', linewidth=0.5)
            self.ax_spectral.minorticks_on()
            if not self.autoscale_y_checkbox.isChecked():
                self.ax_spectral.set_ylim(bottom=-0.05, top=1.05) 
            else:
                self.ax_spectral.autoscale(enable=True, axis='y') 

            if res['l'].size > 1 : 
                self.ax_spectral.set_xlim(res['l'][0], res['l'][-1])
            elif res['l'].size == 1:
                self.ax_spectral.set_xlim(res['l'][0]-1, res['l'][0]+1)


            if spectral_lines: self.ax_spectral.legend()
        
        self.fig_spectral.tight_layout(pad=2.0) 
        self.canvas_spectral.draw()

        if spectral_lines:
            try:
                cursor_objects = mplcursors.cursor(spectral_lines, hover=mplcursors.HoverMode.Transient) 
                
                active_cursors_group = []
                if isinstance(cursor_objects, list):
                    for c_obj in cursor_objects: 
                        c_obj.connect("add", lambda sel: sel.annotation.set_text(f"λ={sel.target[0]:.2f} nm\n{sel.artist.get_label()}={sel.target[1]:.3f}"))
                    active_cursors_group.extend(cursor_objects)
                elif isinstance(cursor_objects, mplcursors.Cursor): 
                    cursor_objects.connect("add", lambda sel: sel.annotation.set_text(f"λ={sel.target[0]:.2f} nm\n{sel.artist.get_label()}={sel.target[1]:.3f}"))
                    active_cursors_group.append(cursor_objects)
                if active_cursors_group:
                    self.cursors_spectral.append(active_cursors_group)

            except Exception as e_cursor:
                print(f"Erreur mplcursors spectral: {e_cursor}")

    def plot_angular_data(self, res, n_super_val):
        self.ax_angular.clear()
        self.clear_plot_cursors(self.cursors_angular)
        
        angular_lines = []
        if res['inc_a'].size > 0 :
            if self.plot_rs_checkbox.isChecked() and res['Rs_a'].size == res['inc_a'].size : 
                l5, = self.ax_angular.plot(res['inc_a'], res['Rs_a'], label='Rs', linestyle='--')
                angular_lines.append(l5)
            if self.plot_rp_checkbox.isChecked() and res['Rp_a'].size == res['inc_a'].size:
                l6, = self.ax_angular.plot(res['inc_a'], res['Rp_a'], label='Rp', linestyle='--')
                angular_lines.append(l6)
            if self.plot_ts_checkbox.isChecked() and res['Ts_a'].size == res['inc_a'].size:
                l7, = self.ax_angular.plot(res['inc_a'], res['Ts_a'], label='Ts', linestyle='-')
                angular_lines.append(l7)
            if self.plot_tp_checkbox.isChecked() and res['Tp_a'].size == res['inc_a'].size:
                l8, = self.ax_angular.plot(res['inc_a'], res['Tp_a'], label='Tp', linestyle='-')
                angular_lines.append(l8)

            self.ax_angular.set_xlabel("Angle d'incidence (degrés)")
            self.ax_angular.set_ylabel('Reflectance / Transmittance')
            title_ang = "Tracé angulaire"
            if res['l_a'].size > 0: 
                 title_ang += f" (λ = {res['l_a'][0]:.0f} nm"
            title_ang += f", n_super={n_super_val:.2f})" if res['l_a'].size > 0 else f"(n_super={n_super_val:.2f})"
            self.ax_angular.set_title(title_ang)

            self.ax_angular.grid(True, which='major', color='grey', linestyle='-', linewidth=0.7)
            self.ax_angular.grid(True, which='minor', color='lightgrey', linestyle=':', linewidth=0.5)
            self.ax_angular.minorticks_on()
            if not self.autoscale_y_checkbox.isChecked():
                self.ax_angular.set_ylim(bottom=-0.05, top=1.05) 
            else:
                self.ax_angular.autoscale(enable=True, axis='y') 

            if res['inc_a'].size > 1: 
                self.ax_angular.set_xlim(res['inc_a'][0], res['inc_a'][-1])
            elif res['inc_a'].size == 1:
                self.ax_angular.set_xlim(res['inc_a'][0]-1, res['inc_a'][0]+1)


            if angular_lines: self.ax_angular.legend()
        
        self.fig_angular.tight_layout(pad=2.0) 
        self.canvas_angular.draw()

        if angular_lines:
            try:
                cursor_objects = mplcursors.cursor(angular_lines, hover=mplcursors.HoverMode.Transient)
                active_cursors_group = []
                if isinstance(cursor_objects, list):
                    for c_obj in cursor_objects: 
                        c_obj.connect("add", lambda sel: sel.annotation.set_text(f"θ={sel.target[0]:.2f}°\n{sel.artist.get_label()}={sel.target[1]:.3f}"))
                    active_cursors_group.extend(cursor_objects)
                elif isinstance(cursor_objects, mplcursors.Cursor):
                    cursor_objects.connect("add", lambda sel: sel.annotation.set_text(f"θ={sel.target[0]:.2f}°\n{sel.artist.get_label()}={sel.target[1]:.3f}"))
                    active_cursors_group.append(cursor_objects)
                if active_cursors_group:
                     self.cursors_angular.append(active_cursors_group)
            except Exception as e_cursor:
                print(f"Erreur mplcursors angulaire: {e_cursor}")

    def plot_stack_visualization(self, ep, n_super_val, nH_r, nH_i, nL_r, nL_i, nSub_r, nSub_i, emp_str_val):
        self.ax_refractive_index_profile.clear()

        if not emp_str_val.strip() or not ep : 
            self.ax_refractive_index_profile.text(0.5, 0.5, "Aucune couche à visualiser", ha='center', va='center', transform=self.ax_refractive_index_profile.transAxes)
            self.canvas_stack_vis.draw()
            return

        indices_complex_layers = [nH_r - 1j * nH_i if i % 2 == 0 else nL_r - 1j * nL_i for i in range(len(emp_str_val.split(',')))]
        n_reel_layers = [np.real(n) for n in indices_complex_layers] 
        n_super_real = float(n_super_val) # Assuming n_super_val is purely real from input
        n_sub_real = float(nSub_r)


        ep_cum = np.cumsum(ep)
        current_ep_cum_max = ep_cum[-1] if ep_cum.size > 0 else 0

        # Construction du profil d'indice : Superstrat -> Couches -> Substrat
        # Avec 'steps-post', on doit construire les coordonnées pour que chaque transition soit correcte
        
        # 1. Superstrat : de -50 à 0 (avant les couches)
        x_coords_profile = [-50, 0] 
        y_coords_profile = [n_super_real, n_super_real]

        # 2. Couches : de 0 à current_ep_cum_max
        # Pour chaque couche, on trace de son début à sa fin
        for i_layer_plot in range(len(n_reel_layers)): 
            layer_start = ep_cum[i_layer_plot-1] if i_layer_plot > 0 else 0
            layer_end = ep_cum[i_layer_plot]
            x_coords_profile.extend([layer_start, layer_end])
            y_coords_profile.extend([n_reel_layers[i_layer_plot], n_reel_layers[i_layer_plot]])
        
        # 3. Substrat : de current_ep_cum_max à current_ep_cum_max + 50 (après toutes les couches)
        x_coords_profile.extend([current_ep_cum_max, current_ep_cum_max + 50])
        y_coords_profile.extend([n_sub_real, n_sub_real]) 

        self.ax_refractive_index_profile.plot(x_coords_profile, y_coords_profile, drawstyle='steps-post', color='darkblue')
        self.ax_refractive_index_profile.set_xlabel('Épaisseur cumulée (nm)')
        self.ax_refractive_index_profile.set_ylabel('Partie réelle de l\'indice')
        self.ax_refractive_index_profile.set_title("Profil d'indice et épaisseur des couches")
        self.ax_refractive_index_profile.grid(True, which='major', color='grey', linestyle='-', linewidth=0.7)
        self.ax_refractive_index_profile.grid(True, which='minor', color='lightgrey', linestyle=':', linewidth=0.5)
        self.ax_refractive_index_profile.minorticks_on()
        self.ax_refractive_index_profile.set_xlim(-50, current_ep_cum_max + 50 if current_ep_cum_max > 0 else 50)
        
        all_n_values_plot = [n_super_real, n_sub_real] + n_reel_layers
        min_n_plot = min(all_n_values_plot) if all_n_values_plot else 0.8 
        max_n_plot = max(all_n_values_plot) if all_n_values_plot else 2.5 
        self.ax_refractive_index_profile.set_ylim(min_n_plot - 0.2, max_n_plot + 0.2)

        # Text for Superstrate and Substrate
        y_text_pos = self.ax_refractive_index_profile.get_ylim()[0] + 0.05
        self.ax_refractive_index_profile.text(-25, y_text_pos, "SUPERSTRAT", ha='center', va='bottom', fontsize=9, color='black')
        self.ax_refractive_index_profile.text(current_ep_cum_max + 25 if current_ep_cum_max > 0 else 25, y_text_pos, "SUBSTRAT", ha='center', va='bottom', fontsize=9, color='black')


        current_pos_label = 0
        for i_label, thickness_label in enumerate(ep): 
            label_x = current_pos_label + thickness_label / 2
            label_y = n_reel_layers[i_label] + 0.05 
            # Adjust label_y if it's too close to plot limits or other indices
            if max_n_plot > min_n_plot: # Avoid division by zero if all indices are same
                if label_y > max_n_plot + 0.15 * (max_n_plot - min_n_plot) : label_y = n_reel_layers[i_label] - 0.1 * (max_n_plot - min_n_plot)
                if label_y < min_n_plot - 0.15 * (max_n_plot - min_n_plot) : label_y = n_reel_layers[i_label] + 0.1 * (max_n_plot - min_n_plot)
            
            self.ax_refractive_index_profile.text(label_x, label_y, f"C{i_label+1}\n{thickness_label:.1f} nm",
                                                  ha='center', va='bottom', fontsize=7, color='red',
                                                  bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.7))
            if i_label < len(ep) -1 : 
                 self.ax_refractive_index_profile.axvline(x=ep_cum[i_label], color='gray', linestyle=':', linewidth=0.8)
            current_pos_label += thickness_label
        
        if ep: # Draw line at the start of the stack (interface superstrate/first layer)
            self.ax_refractive_index_profile.axvline(x=0, color='gray', linestyle=':', linewidth=0.8)

        self.fig_stack_vis.tight_layout(pad=2.0)
        self.canvas_stack_vis.draw()

    def _perform_recalculation_and_plot(self):
        """Effectue le recalcul et le tracé avec gestion robuste des erreurs."""
        self.status_bar.showMessage("Calcul en cours...", 0) 
        QApplication.processEvents() 

        try:
            values = {}
            valid_inputs = True
            for name_cfg, widget_cfg in self.entry_vars_qt.items():
                text_val = widget_cfg.text().strip() 
                validator = widget_cfg.validator()
                if validator:
                    state, _, _ = validator.validate(text_val,0)
                    if state != QValidator.State.Acceptable:
                        widget_cfg.setProperty("validity", "invalid")
                        widget_cfg.style().unpolish(widget_cfg)
                        widget_cfg.style().polish(widget_cfg)
                        valid_inputs = False 
                    else:
                         widget_cfg.setProperty("validity", "valid")
                         widget_cfg.style().unpolish(widget_cfg)
                         widget_cfg.style().polish(widget_cfg)

                if isinstance(validator, QDoubleValidator) or isinstance(validator, QIntValidator):
                    if not text_val: 
                        self.status_bar.showMessage(f"Erreur: Le champ pour '{name_cfg}' ne peut pas être vide.", 5000)
                        widget_cfg.setProperty("validity", "invalid")
                        widget_cfg.style().unpolish(widget_cfg)
                        widget_cfg.style().polish(widget_cfg)
                        valid_inputs = False
                        continue
                    
                    # Conversion robuste avec gestion point/virgule
                    if isinstance(validator, QDoubleValidator):
                        value, success = safe_str_to_float(text_val)
                        if not success:
                            self.status_bar.showMessage(f"Erreur: Valeur numérique invalide pour '{name_cfg}': {text_val}", 5000)
                            widget_cfg.setProperty("validity", "invalid")
                            widget_cfg.style().unpolish(widget_cfg)
                            widget_cfg.style().polish(widget_cfg)
                            valid_inputs = False
                            continue
                        values[name_cfg] = value
                    else:  # QIntValidator
                        value, success = safe_str_to_int(text_val)
                        if not success:
                            self.status_bar.showMessage(f"Erreur: Valeur entière invalide pour '{name_cfg}': {text_val}", 5000)
                            widget_cfg.setProperty("validity", "invalid")
                            widget_cfg.style().unpolish(widget_cfg)
                            widget_cfg.style().polish(widget_cfg)
                            valid_inputs = False
                            continue
                        values[name_cfg] = value
                else: 
                    values[name_cfg] = text_val 
            
            if not valid_inputs:
                self.status_bar.showMessage("Erreur dans les paramètres d'entrée. Veuillez corriger.", 5000)
                return

            # Parsing robuste de l'empilement avec gestion point/virgule
            emp_str_val = values.get('emp_str', '')
            if emp_str_val and emp_str_val.strip():
                emp_factors_list, success, error_msg = parse_empilement_string(emp_str_val)
                if not success:
                    self.status_bar.showMessage(f"Erreur empilement: {error_msg}", 5000)
                    if 'emp_str' in self.entry_vars_qt:
                        emp_entry = self.entry_vars_qt['emp_str']
                        emp_entry.setProperty("validity", "invalid")
                        emp_entry.style().unpolish(emp_entry)
                        emp_entry.style().polish(emp_entry)
                    valid_inputs = False
                    return
                # Vérifier qu'on a au moins une valeur
                if not emp_factors_list:
                    self.status_bar.showMessage("Erreur: L'empilement doit contenir au moins une valeur.", 5000)
                    if 'emp_str' in self.entry_vars_qt:
                        emp_entry = self.entry_vars_qt['emp_str']
                        emp_entry.setProperty("validity", "invalid")
                        emp_entry.style().unpolish(emp_entry)
                        emp_entry.style().polish(emp_entry)
                    valid_inputs = False
                    return
            
            nH = values['nH_r'] - 1j * values['nH_i']
            nL = values['nL_r'] - 1j * values['nL_i']
            nSub = values['nSub_r'] - 1j * values['nSub_i'] # Substrate is now complex
            n_superstrate_val = values['n_super'] # This is purely real from input validator

            substrat_fini_val = self.substrat_fini_checkbox.isChecked() 
            emp_str_val = values.get('emp_str', '')
            
            # Calculer le nombre de couches de manière robuste
            if emp_str_val and emp_str_val.strip():
                emp_factors_for_count, _, _ = parse_empilement_string(emp_str_val)
                num_layers = len(emp_factors_for_count) if emp_factors_for_count else 0
            else:
                num_layers = 0
            
            export_excel = self.export_excel_checkbox.isChecked()

            if values['l_range_deb'] >= values['l_range_fin'] or values['l_step'] <=0:
                self.status_bar.showMessage("Erreur: Intervalle spectral ou pas invalide.", 5000)
                return
            if values['a_range_deb'] >= values['a_range_fin'] or values['a_step'] <=0:
                self.status_bar.showMessage("Erreur: Intervalle angulaire ou pas invalide.", 5000)
                return

            res, ep = calcul_empilement(
                nH, nL, nSub, values['l0'], emp_str_val,
                (values['l_range_deb'], values['l_range_fin']), values['l_step'],
                (values['a_range_deb'], values['a_range_fin']), values['a_step'],
                values['inc'], n_superstrate_val, substrat_fini_val
            )
            
            self.plot_spectral_data(res, values['inc'], n_superstrate_val)
            self.plot_angular_data(res, n_superstrate_val) 
            self.plot_stack_visualization(ep, n_superstrate_val, values['nH_r'], values['nH_i'], 
                                          values['nL_r'], values['nL_i'], values['nSub_r'], values['nSub_i'], emp_str_val)
            self.status_bar.showMessage("Prêt. Calcul et affichage terminés.", 5000)

            if export_excel:
                self.sauvegarder_excel(values, res, substrat_fini_val, num_layers)

        except ValueError as ve: 
            # Erreur de valeur - afficher mais ne pas planter
            try:
                if not self.recalculation_timer.isActive(): 
                    QMessageBox.critical(self, "Erreur de Valeur", str(ve))
                    self.status_bar.showMessage(f"Erreur de valeur: {ve}", 5000)
            except Exception as e_msg:
                # Si même l'affichage d'erreur échoue, juste logger
                self.status_bar.showMessage(f"Erreur de valeur: {ve}", 5000)
        except (TypeError, AttributeError, KeyError) as te:
            # Erreurs de type - ne pas planter
            try:
                if not self.recalculation_timer.isActive():
                    error_msg = f"Erreur de type: {str(te)}"
                    self.status_bar.showMessage(error_msg, 5000)
            except:
                pass  # Ignorer si même l'affichage échoue
        except Exception as e:
            # Toute autre erreur - capturer pour éviter le plantage
            try:
                if not self.recalculation_timer.isActive():
                    error_msg = f"Erreur: {str(e)}"
                    # Ne pas afficher de MessageBox pour les erreurs non critiques
                    self.status_bar.showMessage(error_msg, 5000)
            except:
                pass  # Ignorer complètement si même l'affichage échoue


    def sauvegarder_excel(self, params_entree, res_calcul, substrat_fini_val, num_layers_val):
        self.status_bar.showMessage("Sauvegarde Excel en cours...", 0)
        QApplication.processEvents()
        try:
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y-%m-%d-%H-%M-%S")
            excel_file_default_name = f"Resultats_empilement_{num_layers_val}_couches_{timestamp}.xlsx"
            
            filename, _ = QFileDialog.getSaveFileName(
                self, 
                "Sauvegarder les résultats Excel", 
                excel_file_default_name, 
                "Fichiers Excel (*.xlsx);;Tous les fichiers (*)"
            )

            if not filename: 
                self.status_bar.showMessage("Sauvegarde Excel annulée.", 3000)
                return
            
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                params_for_excel = params_entree.copy()
                # Assurer que tous les paramètres sont bien des types simples pour le DataFrame
                for key, val in params_for_excel.items():
                    if isinstance(val, (np.complex128, np.complex_)):
                        params_for_excel[key] = str(val) # Convertir complexe en string
                    elif isinstance(val, (np.float64, np.int_)):
                        params_for_excel[key] = float(val) # ou int(val)

                df_params = pd.DataFrame.from_dict(params_for_excel, orient='index', columns=['Valeur'])
                df_params.loc['Substrat Fini'] = substrat_fini_val
                df_params.to_excel(writer, sheet_name='Paramètres')

                if res_calcul['l'].size > 0 : 
                    df_spectral_data = {'Longueur d\'onde (nm)': res_calcul['l']}
                    if self.plot_rs_checkbox.isChecked() and 'Rs_s' in res_calcul and res_calcul['Rs_s'].size == res_calcul['l'].size: df_spectral_data['Rs'] = res_calcul['Rs_s']
                    if self.plot_rp_checkbox.isChecked() and 'Rp_s' in res_calcul and res_calcul['Rp_s'].size == res_calcul['l'].size: df_spectral_data['Rp'] = res_calcul['Rp_s']
                    if self.plot_ts_checkbox.isChecked() and 'Ts_s' in res_calcul and res_calcul['Ts_s'].size == res_calcul['l'].size: df_spectral_data['Ts'] = res_calcul['Ts_s']
                    if self.plot_tp_checkbox.isChecked() and 'Tp_s' in res_calcul and res_calcul['Tp_s'].size == res_calcul['l'].size: df_spectral_data['Tp'] = res_calcul['Tp_s']
                    if len(df_spectral_data) > 1: 
                        df_spectral = pd.DataFrame(df_spectral_data)
                        df_spectral.to_excel(writer, sheet_name='Données Spectrales', index=False)


                if res_calcul['inc_a'].size > 0: 
                    df_angular_data = {'Angle (°)': res_calcul['inc_a']}
                    if self.plot_rs_checkbox.isChecked() and 'Rs_a' in res_calcul and res_calcul['Rs_a'].size == res_calcul['inc_a'].size: df_angular_data['Rs'] = res_calcul['Rs_a']
                    if self.plot_rp_checkbox.isChecked() and 'Rp_a' in res_calcul and res_calcul['Rp_a'].size == res_calcul['inc_a'].size: df_angular_data['Rp'] = res_calcul['Rp_a']
                    if self.plot_ts_checkbox.isChecked() and 'Ts_a' in res_calcul and res_calcul['Ts_a'].size == res_calcul['inc_a'].size: df_angular_data['Ts'] = res_calcul['Ts_a']
                    if self.plot_tp_checkbox.isChecked() and 'Tp_a' in res_calcul and res_calcul['Tp_a'].size == res_calcul['inc_a'].size: df_angular_data['Tp'] = res_calcul['Tp_a']
                    if len(df_angular_data) > 1: 
                        df_angular = pd.DataFrame(df_angular_data)
                        df_angular.to_excel(writer, sheet_name='Données Angulaires', index=False)
                
                # Auto-ajustement des largeurs de colonnes
                for sheet_name_excel in writer.sheets: 
                    worksheet = writer.sheets[sheet_name_excel]
                    df_to_format = None
                    if sheet_name_excel == 'Paramètres':
                        df_to_format = df_params # Utiliser le DataFrame original des paramètres
                    elif sheet_name_excel == 'Données Spectrales' and 'df_spectral' in locals():
                        df_to_format = df_spectral
                    elif sheet_name_excel == 'Données Angulaires' and 'df_angular' in locals():
                        df_to_format = df_angular
                    
                    if df_to_format is not None:
                        if sheet_name_excel == 'Paramètres':
                            # Pour la feuille des paramètres, ajuster la colonne de l'index et la colonne 'Valeur'
                            idx_max_len = max(df_to_format.index.astype(str).map(len).max(), len(str(df_to_format.index.name or "Index"))) + 2
                            worksheet.set_column(0, 0, idx_max_len)
                            if 'Valeur' in df_to_format.columns:
                                val_max_len = max(df_to_format['Valeur'].astype(str).map(len).max(), len('Valeur')) + 2
                                worksheet.set_column(1, 1, val_max_len)
                        else: # Pour les autres feuilles
                            for idx_col, col_name_excel in enumerate(df_to_format.columns): 
                                series = df_to_format[col_name_excel]
                                max_val_len = series.astype(str).map(len).max() if not series.empty else 0
                                max_len_col = max(max_val_len, len(str(col_name_excel))) + 2  
                                worksheet.set_column(idx_col, idx_col, max_len_col)
            
            QMessageBox.information(self, "Sauvegarde Réussie", f"Résultats enregistrés dans {filename}")
            self.status_bar.showMessage(f"Résultats exportés vers : {filename}", 5000)

        except Exception as e_excel: 
            QMessageBox.critical(self, "Erreur de Sauvegarde Excel", f"Impossible d'enregistrer le fichier Excel : {e_excel}")
            self.status_bar.showMessage("Erreur lors de la sauvegarde Excel.", 5000)


# Core calculation function
def calcul_empilement(nH, nL, nSub_complex, l0, emp_str, l_range, l_step, a_range, a_step, 
                      inc_deg_in_super, n_superstrate_real, substrat_fini):
    """
    Calcule les propriétés optiques d'un empilement de couches minces.

    Args:
        nH (complex): Indice de réfraction complexe du matériau haut indice.
        nL (complex): Indice de réfraction complexe du matériau bas indice.
        nSub_complex (complex): Indice de réfraction complexe du substrat.
        l0 (float): Longueur d'onde de centrage pour QWOT (nm).
        emp_str (str): Description de l'empilement (ex: "1,0.5,1").
        l_range (tuple): (début, fin) de l'intervalle spectral (nm).
        l_step (float): Pas spectral (nm).
        a_range (tuple): (début, fin) de l'intervalle angulaire (degrés).
        a_step (float): Pas angulaire (degrés).
        inc_deg_in_super (float): Angle d'incidence nominal (design QWOT & tracé spectral) en degrés, dans le superstrat.
        n_superstrate_real (float): Indice de réfraction (réel) du superstrat.
        substrat_fini (bool): True si réflexions multiples sur face arrière du substrat.

    Returns:
        tuple: (dict_resultats, liste_epaisseurs_physiques)
    """
    if l_range[0] >= l_range[1] or l_step <= 0: 
        l_nm = np.array([])
    else:
        l_nm = np.arange(l_range[0], l_range[1] + l_step, l_step)

    if a_range[0] >= a_range[1] or a_step <= 0: 
        theta_inc_ang_deg = np.array([])
    else:
        theta_inc_ang_deg = np.arange(a_range[0], a_range[1] + a_step, a_step)
    
    theta_inc_spectral_rad = np.radians(inc_deg_in_super)
    theta_inc_ang_rad = np.radians(theta_inc_ang_deg)
    l_ang_nm = np.array([l0]) # Pour le tracé angulaire, on utilise la longueur d'onde de centrage
    
    # Calcul des épaisseurs physiques (ep) en tenant compte de l'incidence nominale
    theta_nominal_design_rad = np.radians(inc_deg_in_super)
    
    # Parsing robuste de l'empilement avec gestion point/virgule
    if not emp_str or not emp_str.strip():
        emp_factors = []
        ep_physical_nm = []
    else:
        try:
            emp_factors, parse_success, parse_error = parse_empilement_string(emp_str)
            if not parse_success:
                raise ValueError(f"Erreur parsing empilement: {parse_error}")
            if not emp_factors:
                raise ValueError("L'empilement ne contient aucune valeur valide.") 

            ep_physical_nm = []
            for i, factor_qwot in enumerate(emp_factors):
                n_layer_complex = nH if i % 2 == 0 else nL
                n_layer_real = np.real(n_layer_complex)

                if n_layer_real <= 0:
                    raise ValueError(f"L'indice réel de la couche {i+1} doit être positif. Obtenu: {n_layer_real:.4f}")

                # Loi de Snell: n_super * sin(theta_super) = n_layer * sin(theta_layer)
                # alpha_snell_design = n_superstrate_real * sin(theta_nominal_design_rad)
                # sin_theta_layer_design = alpha_snell_design / n_layer_real
                val_snell_incident_design = n_superstrate_real * np.sin(theta_nominal_design_rad)

                if abs(val_snell_incident_design) > n_layer_real and not np.isclose(abs(val_snell_incident_design), n_layer_real):
                    raise ValueError(
                        f"QWOT impossible pour couche {i+1} (n={n_layer_real:.4f}): "
                        f"Incidence design ({inc_deg_in_super:.2f}° dans n_super={n_superstrate_real:.2f}) "
                        f"mène à RTI (n_couche < n_super * sin(theta_super) => {n_layer_real:.4f} < {abs(val_snell_incident_design):.4f})."
                    )
                
                cos_theta_layer_sq_design = 1.0 - (val_snell_incident_design / n_layer_real)**2
                
                if cos_theta_layer_sq_design < 0:
                    if np.isclose(cos_theta_layer_sq_design, 0): cos_theta_layer_sq_design = 0.0
                    else: raise ValueError(f"Erreur interne QWOT couche {i+1}: cos²(θ_design) < 0.")
                
                cos_theta_layer_design = np.sqrt(cos_theta_layer_sq_design)

                if np.isclose(cos_theta_layer_design, 0.0):
                    if factor_qwot == 0: ep_nm = 0.0
                    else: raise ValueError(f"QWOT impossible pour couche {i+1} (n={n_layer_real:.4f}): Angle critique pour design QWOT.")
                else:
                    ep_nm = (factor_qwot * l0) / (4 * n_layer_real * cos_theta_layer_design)
                ep_physical_nm.append(ep_nm)
        except ValueError as e_val:
            raise ValueError(str(e_val))

    matrices_stockees_layers = {} # Cache pour les matrices de couches individuelles

    def calcul_M_couche(pol, alpha_snell_calc, n_layer_cplx, physical_thickness_nm, lambda_nm_calc):
        """Calcule la matrice caractéristique d'UNE couche."""
        key = (pol, alpha_snell_calc, n_layer_cplx, physical_thickness_nm, lambda_nm_calc)
        if key in matrices_stockees_layers:
            return matrices_stockees_layers[key]

        # Admittance optique de la couche
        eta_layer_sqrt_arg = n_layer_cplx**2 - alpha_snell_calc**2
        eta_layer_sqrt = np.sqrt(eta_layer_sqrt_arg + 0j) # +0j pour gérer les arguments négatifs (ondes évanescentes)
        
        if pol == 'p':
            eta_layer_adm = (n_layer_cplx**2 / eta_layer_sqrt if eta_layer_sqrt != 0 else np.inf)
        else: # 's'
            eta_layer_adm = eta_layer_sqrt
        
        # Phase optique: phi = (2*pi/lambda) * n_couche * d_couche * cos(theta_couche)
        # n_couche * cos(theta_couche) = sqrt(n_couche^2 - (n_super*sin(theta_super))^2) = eta_layer_sqrt
        phi = (2 * np.pi / lambda_nm_calc) * eta_layer_sqrt * physical_thickness_nm
        
        cos_phi = np.cos(phi)
        sin_phi = np.sin(phi)

        if eta_layer_adm == 0 or np.isinf(eta_layer_adm) or np.isnan(eta_layer_adm):
            # Cas limite (ex: angle critique exact, absorption très forte rendant l'admittance nulle ou infinie)
            # La matrice devient une identité ou une forme dégénérée. Pour la propagation, l'identité est plus sûre.
             mat_interface = np.eye(2, dtype=complex)
        else:
            mat_interface = np.array([[cos_phi, (1j / eta_layer_adm) * sin_phi],
                                      [1j * eta_layer_adm * sin_phi, cos_phi]], dtype=complex)
        matrices_stockees_layers[key] = mat_interface
        return mat_interface

    def calcul_RT_globale(longueurs_onde_nm_arr, angles_rad_in_super_arr):
        if not longueurs_onde_nm_arr.size or not angles_rad_in_super_arr.size:
            return np.zeros((0, 0, 4)) # Rs, Rp, Ts, Tp

        RT_results = np.zeros((len(longueurs_onde_nm_arr), len(angles_rad_in_super_arr), 4))

        for i_l, current_l_nm in enumerate(longueurs_onde_nm_arr):
            for i_a, current_theta_rad_super in enumerate(angles_rad_in_super_arr):
                
                alpha_snell = n_superstrate_real * np.sin(current_theta_rad_super)
                if abs(alpha_snell) > n_superstrate_real and not np.isclose(abs(alpha_snell), n_superstrate_real):
                    # Ceci ne devrait pas arriver si current_theta_rad_super est réel. Sécurité.
                    alpha_snell = np.sign(alpha_snell) * n_superstrate_real
                
                for pol_idx, pol_type in enumerate(['s', 'p']):
                    # Matrice globale de l'empilement
                    M_globale = np.eye(2, dtype=complex)
                    if emp_factors: # S'il y a des couches
                        for i_couche, factor_qwt_couche in enumerate(emp_factors):
                            n_cplx_couche = nH if i_couche % 2 == 0 else nL
                            ep_phys_couche = ep_physical_nm[i_couche]
                            M_c = calcul_M_couche(pol_type, alpha_snell, n_cplx_couche, ep_phys_couche, current_l_nm)
                            M_globale = M_c @ M_globale # Multiplication à gauche car on part du substrat vers le superstrat
                    
                    # Admittance du superstrat (milieu incident)
                    eta_super_sqrt_arg = n_superstrate_real**2 - alpha_snell**2
                    eta_super_sqrt = np.sqrt(eta_super_sqrt_arg + 0j)
                    eta_super_adm = (n_superstrate_real**2 / eta_super_sqrt if eta_super_sqrt != 0 else np.inf) if pol_type == 'p' else eta_super_sqrt
                    
                    # Admittance du substrat (milieu émergent pour l'empilement)
                    eta_sub_sqrt_arg = nSub_complex**2 - alpha_snell**2
                    eta_sub_sqrt = np.sqrt(eta_sub_sqrt_arg + 0j)
                    eta_sub_adm = (nSub_complex**2 / eta_sub_sqrt if eta_sub_sqrt != 0 else np.inf) if pol_type == 'p' else eta_sub_sqrt

                    # Calcul de r et t pour l'empilement sur substrat semi-infini
                    # Formules de Macleod (Thin Film Optical Filters, 5th ed., Eq 2.116, 2.117)
                    # Y = C/B, r = (eta_super * B - C) / (eta_super * B + C)
                    # t = 2 * eta_super / (eta_super * B + C)
                    # Où M_globale = [[A, B], [C, D]] (attention, ma définition de M est l'inverse de celle de Macleod parfois)
                    # Ma matrice M va du substrat vers le superstrat.
                    # Si M = [[M00, M01], [M10, M11]]
                    # Dénominateur: den = eta_super_adm * M_globale[0,0] + eta_sub_adm * M_globale[1,1] + eta_super_adm * eta_sub_adm * M_globale[0,1] + M_globale[1,0]
                    # Cette formule est pour une matrice allant du superstrat vers le substrat.
                    # Pour ma matrice (substrat -> superstrat):
                    # M_globale[0,0] = D_macleod, M_globale[0,1] = C_macleod/eta_sub_macleod, M_globale[1,0] = B_macleod*eta_sub_macleod, M_globale[1,1] = A_macleod
                    # Il est plus simple d'utiliser les formules générales:
                    # r = (eta_super_adm * M11 + eta_super_adm * eta_sub_adm * M01 - M10 - eta_sub_adm * M00) /
                    #     (eta_super_adm * M11 + eta_super_adm * eta_sub_adm * M01 + M10 + eta_sub_adm * M00)
                    # où M = [[M00, M01],[M10, M11]] est la matrice de la couche i (M_c dans ma boucle)
                    # Pour la matrice totale M_globale:
                    m00, m01 = M_globale[0,0], M_globale[0,1]
                    m10, m11 = M_globale[1,0], M_globale[1,1]

                    den_rt = (eta_super_adm * m00 + eta_sub_adm * m11 + eta_super_adm * eta_sub_adm * m01 + m10)
                    if den_rt == 0 or np.isinf(den_rt) or np.isnan(den_rt):
                        r_infini = np.inf # Ou 1 si on veut une réflectivité de 1
                        t_infini = 0
                    else:
                        r_infini = ((eta_super_adm * m00 - eta_sub_adm * m11 + eta_super_adm * eta_sub_adm * m01 - m10) / den_rt)
                        t_infini = (2 * eta_super_adm / den_rt)
                    
                    R_infini = np.abs(r_infini)**2 if not (np.isinf(r_infini) or np.isnan(r_infini)) else 1.0
                    
                    # Transmittance: T = Re(eta_sub_adm / eta_super_adm) * |t|^2 si superstrat et substrat non absorbants
                    # Formule générale: T = Re(eta_sub_adm) / Re(eta_super_adm) * |t|^2 si eta_super_adm est réel (non absorbant)
                    # Si superstrat absorbant, la définition de T est plus complexe. On suppose superstrat non absorbant.
                    if np.real(eta_super_adm) != 0 and not (np.isinf(eta_super_adm) or np.isnan(eta_super_adm) or np.isinf(t_infini) or np.isnan(t_infini)):
                        T_infini = (np.real(eta_sub_adm) / np.real(eta_super_adm)) * np.abs(t_infini)**2
                    else:
                        T_infini = 0.0
                    
                    R_val, T_val = R_infini, T_infini

                    if substrat_fini:
                        # Réflexion face arrière du substrat (interface substrat -> superstrat)
                        # r_sub_arriere = (eta_sub_adm - eta_super_adm) / (eta_sub_adm + eta_super_adm)
                        # Rb = |r_sub_arriere|^2
                        den_Rb = (eta_sub_adm + eta_super_adm)
                        if den_Rb == 0 or np.isinf(den_Rb) or np.isnan(den_Rb):
                            Rb = 1.0
                        else:
                            Rb = np.abs((eta_sub_adm - eta_super_adm) / den_Rb)**2
                        
                        # R' est la réflectivité de l'interface substrat/air (ou superstrat)
                        # T_film est T_infini, R_film est R_infini
                        # Formules pour substrat épais (incohérent)
                        # R = R_film + (T_film^2 * Rb) / (1 - R_film * Rb) ; ici R_film est R_interface_avant_film
                        # Il faut R_interface_arriere_film (R'_film dans les notations de Macleod)
                        # Les R_infini et T_infini sont pour le système film+substrat semi-infini.
                        # Les formules de Heavens (ou Born & Wolf) pour substrat fini sont plus adaptées.
                        # R = R_infini + (T_infini * T_infini_prime * Rb_prime) / (1 - R_infini_prime * Rb_prime)
                        # Où R_infini_prime est la réflectivité vue depuis le substrat vers le film.
                        # Plus simple: utiliser les formules pour intensités (valide si substrat > longueur de cohérence)
                        # R_tot = R_avant + (T_avant * T_prime_avant * R_arriere_sub) / (1 - R_prime_avant * R_arriere_sub)
                        # R_avant = R_infini (réflectivité du film sur substrat semi-infini)
                        # T_avant = T_infini (transmittance du film sur substrat semi-infini)
                        # R_arriere_sub = Rb (réflectivité de l'interface substrat/superstrat)
                        # R_prime_avant: réflectivité du film vu depuis le substrat.
                        # T_prime_avant: transmittance du film vu depuis le substrat.
                        # Si le film est non-absorbant, T_prime_avant = T_avant.
                        # Si le film est symétrique optiquement, R_prime_avant = R_avant. Pas général.
                        
                        # On utilise une approximation courante pour substrat non absorbant et film non absorbant:
                        # R = R_f + (T_f^2 * R_b) / (1 - R_f * R_b) (si R_f est la réflectivité de l'interface avant seule)
                        # C'est plus complexe. Utilisons la formule de réflectivité incohérente:
                        # R_effective = R_12 + (T_12 * T_21 * R_23) / (1 - R_21_prime * R_23)
                        # R_12 = R_infini (film sur substrat semi-infini, vu depuis superstrat)
                        # T_12 = T_infini
                        # T_21 = transmittance du film vu depuis substrat. Si non-absorbant: T_21 = T_12 * (n_super/n_sub) si admittances réelles
                        # R_23 = Rb (substrat -> superstrat)
                        # R_21_prime = Réflectivité du film vu depuis substrat.
                        
                        # Formule plus simple et souvent utilisée (approx valide si pas trop d'absorption dans le film):
                        # R_tot = R_infini + (T_infini**2 * Rb) / (1 - R_infini * Rb)
                        # T_tot = T_infini * (1-Rb) / (1 - R_infini * Rb)
                        # Ces formules supposent que T_film_direct = T_film_inverse (ou compensé par les indices)
                        # et R_film_direct = R_film_inverse.
                        
                        # Vérifions la conservation d'énergie pour T_infini_prime
                        # T_film_prime = T_infini * (Re(eta_super_adm) / Re(eta_sub_adm)) si film non absorbant.
                        # Si on suppose que le film est "réciproque" en termes de transmission d'intensité.
                        
                        den_sub_fini = (1.0 - R_infini * Rb) # Attention R_infini ici est la réflectivité du film sur substrat semi-infini
                                                            # et non la réflectivité de l'interface arrière du film.
                        if den_sub_fini != 0 and not (np.isinf(den_sub_fini) or np.isnan(den_sub_fini)):
                            R_val = R_infini + (T_infini**2 * Rb) / den_sub_fini
                            T_val = T_infini * (1.0 - Rb) / den_sub_fini # (1-Rb) est comme T_sub_super
                        else: # Si dénominateur nul, on garde les valeurs pour substrat infini
                            R_val = R_infini
                            T_val = T_infini
                            
                    # Assurer que les valeurs sont entre 0 et 1 (ou un peu plus pour erreurs numériques)
                    R_val = np.clip(R_val, 0, 1.001)
                    T_val = np.clip(T_val, 0, 1.001)

                    if pol_type == 's':
                        RT_results[i_l, i_a, 0] = R_val # Rs
                        RT_results[i_l, i_a, 2] = T_val # Ts
                    else: # 'p'
                        RT_results[i_l, i_a, 1] = R_val # Rp
                        RT_results[i_l, i_a, 3] = T_val # Tp
        return RT_results

    # Calcul pour le tracé spectral
    RT_spectral = calcul_RT_globale(l_nm, np.array([theta_inc_spectral_rad]))
    
    # Calcul pour le tracé angulaire
    RT_angular = calcul_RT_globale(l_ang_nm, theta_inc_ang_rad)
    
    # Extraction des résultats
    Rs_s_data = RT_spectral[:,0,0] if RT_spectral.size and RT_spectral.shape[0] > 0 and RT_spectral.shape[1] > 0 else np.array([])
    Rp_s_data = RT_spectral[:,0,1] if RT_spectral.size and RT_spectral.shape[0] > 0 and RT_spectral.shape[1] > 0 else np.array([])
    Ts_s_data = RT_spectral[:,0,2] if RT_spectral.size and RT_spectral.shape[0] > 0 and RT_spectral.shape[1] > 0 else np.array([])
    Tp_s_data = RT_spectral[:,0,3] if RT_spectral.size and RT_spectral.shape[0] > 0 and RT_spectral.shape[1] > 0 else np.array([])

    Rs_a_data = RT_angular[0,:,0] if RT_angular.size and RT_angular.shape[0] > 0 and RT_angular.shape[1] > 0 else np.array([])
    Rp_a_data = RT_angular[0,:,1] if RT_angular.size and RT_angular.shape[0] > 0 and RT_angular.shape[1] > 0 else np.array([])
    Ts_a_data = RT_angular[0,:,2] if RT_angular.size and RT_angular.shape[0] > 0 and RT_angular.shape[1] > 0 else np.array([])
    Tp_a_data = RT_angular[0,:,3] if RT_angular.size and RT_angular.shape[0] > 0 and RT_angular.shape[1] > 0 else np.array([])

    return {'l': l_nm, 'inc_spectral_deg': np.array([inc_deg_in_super]), 
            'Rs_s': Rs_s_data, 'Rp_s': Rp_s_data, 'Ts_s': Ts_s_data, 'Tp_s': Tp_s_data, 
            'l_a': l_ang_nm, 'inc_a': theta_inc_ang_deg,
            'Rs_a': Rs_a_data, 'Rp_a': Rp_a_data, 'Ts_a': Ts_a_data, 'Tp_a': Tp_a_data}, ep_physical_nm


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # INPUT_CONFIGS est maintenant défini globalement en haut du fichier.
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec())

