"""
Application Streamlit pour le calcul d'empilement de couches minces
Version ultra ergonomique avec interface moderne
"""

import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from io import BytesIO
from datetime import datetime
import sys
import os

# Configuration de la page
st.set_page_config(
    page_title="Calcul Couches Minces",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Styles CSS personnalisés pour une interface moderne
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #6c7ae0;
        text-align: center;
        padding: 1rem 0;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: 600;
        color: #495057;
        border-bottom: 2px solid #6c7ae0;
        padding-bottom: 0.5rem;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(90deg, #6c7ae0 0%, #5568d3 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #7d8aef 0%, #6578e4 100%);
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(108, 122, 224, 0.3);
    }
    .metric-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #e8ecef 100%);
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #6c7ae0;
        margin: 0.5rem 0;
    }
    .success-box {
        padding: 1rem;
        border-radius: 8px;
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        border-radius: 8px;
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Import des fonctions de calcul depuis le fichier original
# On essaie d'importer directement - PyQt6 doit être installé (même si on ne l'utilise pas)
try:
    from cm_simple7 import calcul_empilement, parse_empilement_string, safe_str_to_float, safe_str_to_int
except ImportError as e:
    st.error(f"❌ Erreur d'import: {str(e)}\n\nAssurez-vous que PyQt6 est installé: pip install PyQt6")
    st.stop()

# Fonctions utilitaires (déjà importées, mais définies aussi localement pour compatibilité)
def safe_str_to_float(text):
    """Convertit une chaîne en float de manière robuste, gérant point et virgule."""
    if not text or not isinstance(text, str):
        return 0.0, False
    text = text.strip()
    if not text:
        return 0.0, False
    try:
        normalized = text.replace(',', '.').replace(' ', '')
        if normalized.count('.') > 1:
            return 0.0, False
        value = float(normalized)
        return value, True
    except (ValueError, AttributeError, TypeError):
        return 0.0, False

def safe_str_to_int(text):
    """Convertit une chaîne en int de manière robuste."""
    if not text or not isinstance(text, str):
        return 0, False
    text = text.strip()
    if not text:
        return 0, False
    try:
        normalized = text.replace(',', '.').replace(' ', '')
        if '.' in normalized:
            float_val = float(normalized)
            return int(round(float_val)), True
        value = int(normalized)
        return value, True
    except (ValueError, AttributeError, TypeError, OverflowError):
        return 0, False

def parse_empilement_string(emp_str):
    """Parse une chaîne d'empilement robuste, gérant point et virgule."""
    if not emp_str or not isinstance(emp_str, str):
        return [], True, ""
    emp_str = emp_str.strip()
    if not emp_str:
        return [], True, ""
    try:
        parts = emp_str.split(',')
        emp_factors = []
        for i, part in enumerate(parts):
            part = part.strip()
            if not part:
                continue
            value, success = safe_str_to_float(part)
            if not success:
                return [], False, f"Valeur invalide à la position {i+1}: '{part}'"
            if value < 0:
                return [], False, f"Valeur négative non autorisée à la position {i+1}: {value}"
            emp_factors.append(value)
        return emp_factors, True, ""
    except Exception as e:
        return [], False, f"Erreur lors du parsing: {str(e)}"

# Supprimer les warnings matplotlib pour Streamlit
import warnings
warnings.filterwarnings('ignore')

# Initialisation du session state pour UNDO/REDO
if 'undo_history' not in st.session_state:
    st.session_state.undo_history = []
if 'redo_history' not in st.session_state:
    st.session_state.redo_history = []
if 'max_undo_steps' not in st.session_state:
    st.session_state.max_undo_steps = 5
if 'current_state' not in st.session_state:
    st.session_state.current_state = {}

def save_current_state():
    """Sauvegarde l'état actuel dans l'historique UNDO."""
    if len(st.session_state.undo_history) >= st.session_state.max_undo_steps:
        st.session_state.undo_history.pop(0)
    st.session_state.undo_history.append(st.session_state.current_state.copy())
    st.session_state.redo_history.clear()

def undo_action():
    """Annule la dernière action."""
    if len(st.session_state.undo_history) <= 1:
        return
    st.session_state.redo_history.append(st.session_state.current_state.copy())
    st.session_state.undo_history.pop()
    if st.session_state.undo_history:
        st.session_state.current_state = st.session_state.undo_history[-1].copy()
        st.rerun()

def redo_action():
    """Refait la dernière action annulée."""
    if not st.session_state.redo_history:
        return
    st.session_state.undo_history.append(st.session_state.current_state.copy())
    st.session_state.current_state = st.session_state.redo_history.pop().copy()
    st.rerun()

# En-tête principal
st.markdown('<div class="main-header">🔬 Calcul d\'Empilement de Couches Minces</div>', unsafe_allow_html=True)

# Sidebar avec les paramètres
with st.sidebar:
    st.markdown("### ⚙️ Paramètres")
    
    # Boutons UNDO/REDO
    col_undo1, col_redo1 = st.columns(2)
    with col_undo1:
        undo_disabled = len(st.session_state.undo_history) <= 1
        if st.button("↶ UNDO", disabled=undo_disabled, help="Annuler la dernière modification"):
            undo_action()
    with col_redo1:
        redo_disabled = len(st.session_state.redo_history) == 0
        if st.button("↷ REDO", disabled=redo_disabled, help="Refaire la dernière annulation"):
            redo_action()
    
    st.markdown("---")
    
    # Section Matériaux
    st.markdown("#### 📦 Matériaux")
    
    n_super = st.number_input(
        "Indice du Superstrat (milieu incident)",
        min_value=0.1, max_value=10.0, value=1.0, step=0.01, format="%.4f",
        help="Indice de réfraction réel du milieu d'où la lumière est incidente"
    )
    
    nH_r = st.number_input(
        "Matériau H (réel)",
        min_value=0.1, max_value=10.0, value=2.25, step=0.01, format="%.4f",
        help="Indice de réfraction réel du matériau haut indice"
    )
    
    nH_i = st.number_input(
        "Matériau H (imaginaire)",
        min_value=0.0, max_value=1.0, value=0.0001, step=0.00001, format="%.5f",
        help="Partie imaginaire de l'indice de H (absorption)"
    )
    
    nL_r = st.number_input(
        "Matériau L (réel)",
        min_value=0.1, max_value=10.0, value=1.48, step=0.01, format="%.4f",
        help="Indice de réfraction réel du matériau bas indice"
    )
    
    nL_i = st.number_input(
        "Matériau L (imaginaire)",
        min_value=0.0, max_value=1.0, value=0.0001, step=0.00001, format="%.5f",
        help="Partie imaginaire de l'indice de L (absorption)"
    )
    
    nSub_r = st.number_input(
        "Substrat (indice réel)",
        min_value=0.1, max_value=10.0, value=1.52, step=0.01, format="%.4f",
        help="Indice de réfraction réel du substrat"
    )
    
    nSub_i = st.number_input(
        "Substrat (indice imaginaire)",
        min_value=0.0, max_value=1.0, value=0.0, step=0.00001, format="%.5f",
        help="Partie imaginaire de l'indice du substrat"
    )
    
    st.markdown("---")
    
    # Section Configuration de l'Empilement
    st.markdown("#### 🔬 Configuration de l'Empilement")
    
    l0 = st.number_input(
        "Longueur d'onde de centrage (nm)",
        min_value=1.0, max_value=2000.0, value=550.0, step=1.0,
        help="Longueur d'onde pour laquelle les épaisseurs QWOT sont calculées"
    )
    
    emp_str = st.text_input(
        "Empilement (QWOT, ex: 1,0.5,1)",
        value="1,1,1,1,1,2,1,1,1,1,1",
        help="Séquence des couches en multiples de QWOT, séparées par des virgules"
    )
    
    # Calcul et affichage du nombre de couches
    emp_factors_check, success_check, _ = parse_empilement_string(emp_str)
    num_layers = len(emp_factors_check) if success_check else 0
    st.metric("Nombre de couches", num_layers)
    
    st.markdown("---")
    
    # Section Paramètres Spectraux
    st.markdown("#### 📊 Paramètres Spectraux")
    
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        l_range_deb = st.number_input(
            "Début (nm)", min_value=1.0, max_value=2000.0, value=400.0, step=1.0
        )
    with col_l2:
        l_range_fin = st.number_input(
            "Fin (nm)", min_value=1.0, max_value=2000.0, value=700.0, step=1.0
        )
    
    l_step = st.number_input(
        "Pas spectral (nm)", min_value=0.01, max_value=100.0, value=1.0, step=0.1, format="%.2f"
    )
    
    st.markdown("---")
    
    # Section Paramètres Angulaires
    st.markdown("#### 📐 Paramètres Angulaires")
    
    inc = st.number_input(
        "Incidence nominale (°)",
        min_value=0.0, max_value=89.99, value=0.0, step=0.1, format="%.2f"
    )
    
    col_a1, col_a2 = st.columns(2)
    with col_a1:
        a_range_deb = st.number_input(
            "Début (°)", min_value=0.0, max_value=89.99, value=0.0, step=0.1, format="%.2f"
        )
    with col_a2:
        a_range_fin = st.number_input(
            "Fin (°)", min_value=0.0, max_value=89.99, value=89.0, step=0.1, format="%.2f"
        )
    
    a_step = st.number_input(
        "Pas angulaire (°)", min_value=0.01, max_value=90.0, value=1.0, step=0.1, format="%.2f"
    )
    
    st.markdown("---")
    
    # Options de Calcul
    st.markdown("#### ⚡ Options")
    
    plot_rs = st.checkbox("Afficher Rs", value=True)
    plot_rp = st.checkbox("Afficher Rp", value=True)
    plot_ts = st.checkbox("Afficher Ts", value=True)
    plot_tp = st.checkbox("Afficher Tp", value=True)
    autoscale_y = st.checkbox("Échelle Y Automatique", value=False)
    substrat_fini = st.checkbox("Substrat fini (réflexions multiples)", value=False)
    export_excel = st.checkbox("Exporter vers Excel", value=False)
    
    st.markdown("---")
    
    # Bouton de calcul
    if st.button("🔄 Calculer", type="primary", use_container_width=True):
        st.session_state.calculate = True
    
    # Bouton réinitialiser
    if st.button("🔄 Réinitialiser", use_container_width=True):
        st.session_state.calculate = False
        # Réinitialisation des valeurs par défaut
        st.rerun()

# Zone principale avec les graphiques
if st.session_state.get('calculate', False):
    try:
        # Validation des paramètres
        if l_range_deb >= l_range_fin or l_step <= 0:
            st.error("❌ Erreur: Intervalle spectral ou pas invalide.")
        elif a_range_deb >= a_range_fin or a_step <= 0:
            st.error("❌ Erreur: Intervalle angulaire ou pas invalide.")
        else:
            # Parsing de l'empilement
            emp_factors, success, error_msg = parse_empilement_string(emp_str)
            if not success:
                st.error(f"❌ Erreur empilement: {error_msg}")
            elif not emp_factors:
                st.error("❌ Erreur: L'empilement doit contenir au moins une valeur.")
            else:
                # Calcul avec barre de progression
                progress_bar = st.progress(0)
                status_text = st.empty()
                status_text.text("⚙️ Calcul en cours...")
                
                # Préparation des paramètres complexes
                nH = nH_r - 1j * nH_i
                nL = nL_r - 1j * nL_i
                nSub = nSub_r - 1j * nSub_i
                
                progress_bar.progress(20)
                
                # Appel de la fonction de calcul
                try:
                    res, ep = calcul_empilement(
                        nH, nL, nSub, l0, emp_str,
                        (l_range_deb, l_range_fin), l_step,
                        (a_range_deb, a_range_fin), a_step,
                        inc, n_super, substrat_fini
                    )
                    
                    progress_bar.progress(80)
                    status_text.text("✅ Calcul terminé! Génération des graphiques...")
                    
                    # Onglets pour les graphiques
                    tab1, tab2, tab3 = st.tabs(["📊 Graphique Spectral", "📐 Graphique Angulaire", "🔬 Visualisation Empilement"])
                    
                    with tab1:
                        # Graphique spectral
                        fig_spectral = plt.figure(figsize=(10, 6))
                        ax_spectral = fig_spectral.add_subplot(111)
                        
                        if res['l'].size > 0:
                            if plot_rs and res['Rs_s'].size == res['l'].size:
                                ax_spectral.plot(res['l'], res['Rs_s'], label='Rs', linestyle='-', linewidth=2)
                            if plot_rp and res['Rp_s'].size == res['l'].size:
                                ax_spectral.plot(res['l'], res['Rp_s'], label='Rp', linestyle='--', linewidth=2)
                            if plot_ts and res['Ts_s'].size == res['l'].size:
                                ax_spectral.plot(res['l'], res['Ts_s'], label='Ts', linestyle='-', linewidth=2)
                            if plot_tp and res['Tp_s'].size == res['l'].size:
                                ax_spectral.plot(res['l'], res['Tp_s'], label='Tp', linestyle='--', linewidth=2)
                            
                            ax_spectral.set_xlabel('Longueur d\'onde (nm)', fontsize=12)
                            ax_spectral.set_ylabel('Reflectance / Transmittance', fontsize=12)
                            ax_spectral.set_title(f"Tracé spectral (n_super={n_super:.2f}, incidence {inc:.1f}°)", fontsize=14, fontweight='bold')
                            ax_spectral.grid(True, which='major', color='grey', linestyle='-', linewidth=0.7, alpha=0.5)
                            ax_spectral.grid(True, which='minor', color='lightgrey', linestyle=':', linewidth=0.5, alpha=0.3)
                            ax_spectral.minorticks_on()
                            
                            if not autoscale_y:
                                ax_spectral.set_ylim(bottom=-0.05, top=1.05)
                            else:
                                ax_spectral.autoscale(enable=True, axis='y')
                            
                            if res['l'].size > 1:
                                ax_spectral.set_xlim(res['l'][0], res['l'][-1])
                            elif res['l'].size == 1:
                                ax_spectral.set_xlim(res['l'][0]-1, res['l'][0]+1)
                            
                            ax_spectral.legend(loc='best', fontsize=10)
                        
                        plt.tight_layout()
                        st.pyplot(fig_spectral)
                        
                        # Bouton de téléchargement
                        buf = BytesIO()
                        fig_spectral.savefig(buf, format='png', dpi=150, bbox_inches='tight')
                        buf.seek(0)
                        st.download_button("💾 Télécharger le graphique spectral", buf.getvalue(), 
                                         "graphique_spectral.png", "image/png")
                    
                    with tab2:
                        # Graphique angulaire
                        fig_angular = plt.figure(figsize=(10, 6))
                        ax_angular = fig_angular.add_subplot(111)
                        
                        if res['inc_a'].size > 0:
                            if plot_rs and res['Rs_a'].size == res['inc_a'].size:
                                ax_angular.plot(res['inc_a'], res['Rs_a'], label='Rs', linestyle='--', linewidth=2)
                            if plot_rp and res['Rp_a'].size == res['inc_a'].size:
                                ax_angular.plot(res['inc_a'], res['Rp_a'], label='Rp', linestyle='--', linewidth=2)
                            if plot_ts and res['Ts_a'].size == res['inc_a'].size:
                                ax_angular.plot(res['inc_a'], res['Ts_a'], label='Ts', linestyle='-', linewidth=2)
                            if plot_tp and res['Tp_a'].size == res['inc_a'].size:
                                ax_angular.plot(res['inc_a'], res['Tp_a'], label='Tp', linestyle='-', linewidth=2)
                            
                            ax_angular.set_xlabel("Angle d'incidence (degrés)", fontsize=12)
                            ax_angular.set_ylabel('Reflectance / Transmittance', fontsize=12)
                            title_ang = "Tracé angulaire"
                            if res['l_a'].size > 0:
                                title_ang += f" (λ = {res['l_a'][0]:.0f} nm"
                            title_ang += f", n_super={n_super:.2f})" if res['l_a'].size > 0 else f"(n_super={n_super:.2f})"
                            ax_angular.set_title(title_ang, fontsize=14, fontweight='bold')
                            ax_angular.grid(True, which='major', color='grey', linestyle='-', linewidth=0.7, alpha=0.5)
                            ax_angular.grid(True, which='minor', color='lightgrey', linestyle=':', linewidth=0.5, alpha=0.3)
                            ax_angular.minorticks_on()
                            
                            if not autoscale_y:
                                ax_angular.set_ylim(bottom=-0.05, top=1.05)
                            else:
                                ax_angular.autoscale(enable=True, axis='y')
                            
                            if res['inc_a'].size > 1:
                                ax_angular.set_xlim(res['inc_a'][0], res['inc_a'][-1])
                            elif res['inc_a'].size == 1:
                                ax_angular.set_xlim(res['inc_a'][0]-1, res['inc_a'][0]+1)
                            
                            ax_angular.legend(loc='best', fontsize=10)
                        
                        plt.tight_layout()
                        st.pyplot(fig_angular)
                        
                        # Bouton de téléchargement
                        buf2 = BytesIO()
                        fig_angular.savefig(buf2, format='png', dpi=150, bbox_inches='tight')
                        buf2.seek(0)
                        st.download_button("💾 Télécharger le graphique angulaire", buf2.getvalue(), 
                                         "graphique_angulaire.png", "image/png")
                    
                    with tab3:
                        # Visualisation de l'empilement
                        fig_stack = plt.figure(figsize=(12, 6))
                        ax_stack = fig_stack.add_subplot(111)
                        
                        if emp_str.strip() and ep:
                            indices_complex_layers = [nH_r - 1j * nH_i if i % 2 == 0 else nL_r - 1j * nL_i 
                                                      for i in range(len(emp_factors))]
                            n_reel_layers = [np.real(n) for n in indices_complex_layers]
                            
                            ep_cum = np.cumsum(ep)
                            current_ep_cum_max = ep_cum[-1] if ep_cum.size > 0 else 0
                            
                            # Profil d'indice : Superstrat -> Couches -> Substrat
                            # 1. Superstrat : de -50 à 0 (avant les couches)
                            x_coords = [-50, 0]
                            y_coords = [n_super, n_super]
                            
                            # 2. Couches : de 0 à current_ep_cum_max
                            # Pour chaque couche, on trace de son début à sa fin
                            for i_layer in range(len(n_reel_layers)):
                                layer_start = ep_cum[i_layer-1] if i_layer > 0 else 0
                                layer_end = ep_cum[i_layer]
                                x_coords.extend([layer_start, layer_end])
                                y_coords.extend([n_reel_layers[i_layer], n_reel_layers[i_layer]])
                            
                            # 3. Substrat : de current_ep_cum_max à current_ep_cum_max + 50 (après toutes les couches)
                            x_coords.extend([current_ep_cum_max, current_ep_cum_max + 50])
                            y_coords.extend([nSub_r, nSub_r])
                            
                            ax_stack.plot(x_coords, y_coords, drawstyle='steps-post', color='darkblue', linewidth=2)
                            ax_stack.set_xlabel('Épaisseur cumulée (nm)', fontsize=12)
                            ax_stack.set_ylabel('Partie réelle de l\'indice', fontsize=12)
                            ax_stack.set_title("Profil d'indice et épaisseur des couches", fontsize=14, fontweight='bold')
                            ax_stack.grid(True, which='major', color='grey', linestyle='-', linewidth=0.7, alpha=0.5)
                            ax_stack.grid(True, which='minor', color='lightgrey', linestyle=':', linewidth=0.5, alpha=0.3)
                            ax_stack.minorticks_on()
                            ax_stack.set_xlim(-50, current_ep_cum_max + 50 if current_ep_cum_max > 0 else 50)
                            
                            all_n_values = [n_super, nSub_r] + n_reel_layers
                            min_n = min(all_n_values) if all_n_values else 0.8
                            max_n = max(all_n_values) if all_n_values else 2.5
                            ax_stack.set_ylim(min_n - 0.2, max_n + 0.2)
                            
                            # Labels
                            y_text_pos = ax_stack.get_ylim()[0] + 0.05
                            ax_stack.text(-25, y_text_pos, "SUPERSTRAT", ha='center', va='bottom', fontsize=10, color='black', fontweight='bold')
                            ax_stack.text(current_ep_cum_max + 25 if current_ep_cum_max > 0 else 25, y_text_pos, "SUBSTRAT", 
                                         ha='center', va='bottom', fontsize=10, color='black', fontweight='bold')
                            
                            current_pos = 0
                            for i_label, thickness in enumerate(ep):
                                label_x = current_pos + thickness / 2
                                label_y = n_reel_layers[i_label] + 0.05
                                if max_n > min_n:
                                    if label_y > max_n + 0.15 * (max_n - min_n):
                                        label_y = n_reel_layers[i_label] - 0.1 * (max_n - min_n)
                                    if label_y < min_n - 0.15 * (max_n - min_n):
                                        label_y = n_reel_layers[i_label] + 0.1 * (max_n - min_n)
                                
                                ax_stack.text(label_x, label_y, f"C{i_label+1}\n{thickness:.1f} nm",
                                            ha='center', va='bottom', fontsize=9, color='red', fontweight='bold',
                                            bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7))
                                if i_label < len(ep) - 1:
                                    ax_stack.axvline(x=ep_cum[i_label], color='gray', linestyle=':', linewidth=1)
                                current_pos += thickness
                            
                            if ep:
                                ax_stack.axvline(x=0, color='gray', linestyle=':', linewidth=1)
                        else:
                            ax_stack.text(0.5, 0.5, "Aucune couche à visualiser", ha='center', va='center', 
                                        transform=ax_stack.transAxes, fontsize=14)
                        
                        plt.tight_layout()
                        st.pyplot(fig_stack)
                        
                        # Bouton de téléchargement
                        buf3 = BytesIO()
                        fig_stack.savefig(buf3, format='png', dpi=150, bbox_inches='tight')
                        buf3.seek(0)
                        st.download_button("💾 Télécharger la visualisation empilement", buf3.getvalue(), 
                                         "visualisation_empilement.png", "image/png")
                    
                    progress_bar.progress(100)
                    status_text.text("✅ Calcul et affichage terminés!")
                    
                    # Export Excel si demandé
                    if export_excel:
                        try:
                            status_text.text("💾 Export Excel en cours...")
                            timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
                            
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                # Paramètres
                                params_dict = {
                                    'n_super': n_super,
                                    'nH_r': nH_r, 'nH_i': nH_i,
                                    'nL_r': nL_r, 'nL_i': nL_i,
                                    'nSub_r': nSub_r, 'nSub_i': nSub_i,
                                    'l0': l0, 'emp_str': emp_str,
                                    'l_range_deb': l_range_deb, 'l_range_fin': l_range_fin, 'l_step': l_step,
                                    'inc': inc, 'a_range_deb': a_range_deb, 'a_range_fin': a_range_fin, 'a_step': a_step,
                                    'Substrat Fini': substrat_fini
                                }
                                df_params = pd.DataFrame.from_dict(params_dict, orient='index', columns=['Valeur'])
                                df_params.to_excel(writer, sheet_name='Paramètres')
                                
                                # Données spectrales
                                if res['l'].size > 0:
                                    df_spectral_data = {'Longueur d\'onde (nm)': res['l']}
                                    if plot_rs and res['Rs_s'].size == res['l'].size:
                                        df_spectral_data['Rs'] = res['Rs_s']
                                    if plot_rp and res['Rp_s'].size == res['l'].size:
                                        df_spectral_data['Rp'] = res['Rp_s']
                                    if plot_ts and res['Ts_s'].size == res['l'].size:
                                        df_spectral_data['Ts'] = res['Ts_s']
                                    if plot_tp and res['Tp_s'].size == res['l'].size:
                                        df_spectral_data['Tp'] = res['Tp_s']
                                    if len(df_spectral_data) > 1:
                                        df_spectral = pd.DataFrame(df_spectral_data)
                                        df_spectral.to_excel(writer, sheet_name='Données Spectrales', index=False)
                                
                                # Données angulaires
                                if res['inc_a'].size > 0:
                                    df_angular_data = {'Angle (°)': res['inc_a']}
                                    if plot_rs and res['Rs_a'].size == res['inc_a'].size:
                                        df_angular_data['Rs'] = res['Rs_a']
                                    if plot_rp and res['Rp_a'].size == res['inc_a'].size:
                                        df_angular_data['Rp'] = res['Rp_a']
                                    if plot_ts and res['Ts_a'].size == res['inc_a'].size:
                                        df_angular_data['Ts'] = res['Ts_a']
                                    if plot_tp and res['Tp_a'].size == res['inc_a'].size:
                                        df_angular_data['Tp'] = res['Tp_a']
                                    if len(df_angular_data) > 1:
                                        df_angular = pd.DataFrame(df_angular_data)
                                        df_angular.to_excel(writer, sheet_name='Données Angulaires', index=False)
                            
                            excel_buffer.seek(0)
                            excel_filename = f"Resultats_empilement_{num_layers}_couches_{timestamp}.xlsx"
                            st.download_button("📊 Télécharger les résultats Excel", excel_buffer.getvalue(), 
                                             excel_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            status_text.text("✅ Export Excel prêt!")
                        except Exception as e_excel:
                            st.error(f"❌ Erreur lors de l'export Excel: {str(e_excel)}")
                    
                    progress_bar.empty()
                    
                except (NotImplementedError, NameError) as ne:
                    st.error(f"❌ La fonction calcul_empilement n'est pas disponible: {str(ne)}")
                    st.info("Assurez-vous que cm_simple7.py est dans le même dossier et que PyQt6 est installé.")
                except Exception as calc_error:
                    st.error(f"❌ Erreur lors du calcul: {str(calc_error)}")
                    st.exception(calc_error)
                
    except Exception as e:
        st.error(f"❌ Erreur lors du calcul: {str(e)}")
        st.exception(e)
else:
    st.info("👈 Configurez les paramètres dans la barre latérale et cliquez sur 'Calculer' pour commencer.")

# Note pour l'utilisateur
st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 Info")
st.sidebar.success(
    "Application Streamlit pour le calcul d'empilement de couches minces.\n\n"
    "Version ultra ergonomique avec interface moderne."
)

