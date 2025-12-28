# Fichier: cm_calc.py
import numpy as np

# --- Fonctions Utilitaires ---

def safe_str_to_float(text):
    """Convertit une chaîne en float de manière robuste."""
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
    """Parse une chaîne d'empilement robuste."""
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
        return [], False, f"Erreur lors du parsing de l'empilement: {str(e)}"

# --- Cœur de Calcul ---

def calcul_empilement(nH, nL, nSub_complex, l0, emp_str, l_range, l_step, a_range, a_step, 
                      inc_deg_in_super, n_superstrate_real, substrat_fini):
    """
    Calcule les propriétés optiques d'un empilement de couches minces.
    Version nettoyée pour usage Web/Backend sans dépendances GUI.
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
    l_ang_nm = np.array([l0]) 
    
    # Parsing
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
            theta_nominal_design_rad = np.radians(inc_deg_in_super)

            for i, factor_qwot in enumerate(emp_factors):
                n_layer_complex = nH if i % 2 == 0 else nL
                n_layer_real = np.real(n_layer_complex)

                if n_layer_real <= 0:
                    raise ValueError(f"L'indice réel de la couche {i+1} doit être positif.")

                val_snell_incident_design = n_superstrate_real * np.sin(theta_nominal_design_rad)

                if abs(val_snell_incident_design) > n_layer_real and not np.isclose(abs(val_snell_incident_design), n_layer_real):
                    raise ValueError(f"QWOT impossible pour couche {i+1}: RTI atteinte.")
                
                cos_theta_layer_sq_design = 1.0 - (val_snell_incident_design / n_layer_real)**2
                cos_theta_layer_design = np.sqrt(max(0, cos_theta_layer_sq_design))

                if np.isclose(cos_theta_layer_design, 0.0):
                    if factor_qwot == 0: ep_nm = 0.0
                    else: raise ValueError(f"QWOT impossible pour couche {i+1}: Angle critique.")
                else:
                    ep_nm = (factor_qwot * l0) / (4 * n_layer_real * cos_theta_layer_design)
                ep_physical_nm.append(ep_nm)
        except ValueError as e_val:
            raise ValueError(str(e_val))

    # --- Logique Matrice Caractéristique ---
    # Note: On recalcule à chaque fois pour éviter les problèmes de cache global dans un env multi-user comme Streamlit
    
    def calcul_M_couche(pol, alpha_snell_calc, n_layer_cplx, physical_thickness_nm, lambda_nm_calc):
        eta_layer_sqrt_arg = n_layer_cplx**2 - alpha_snell_calc**2
        eta_layer_sqrt = np.sqrt(eta_layer_sqrt_arg + 0j)
        
        if pol == 'p':
            eta_layer_adm = (n_layer_cplx**2 / eta_layer_sqrt if eta_layer_sqrt != 0 else np.inf)
        else: # 's'
            eta_layer_adm = eta_layer_sqrt
        
        phi = (2 * np.pi / lambda_nm_calc) * eta_layer_sqrt * physical_thickness_nm
        cos_phi = np.cos(phi)
        sin_phi = np.sin(phi)

        if eta_layer_adm == 0 or np.isinf(eta_layer_adm) or np.isnan(eta_layer_adm):
             mat_interface = np.eye(2, dtype=complex)
        else:
            mat_interface = np.array([[cos_phi, (1j / eta_layer_adm) * sin_phi],
                                      [1j * eta_layer_adm * sin_phi, cos_phi]], dtype=complex)
        return mat_interface

    def calcul_RT_globale(longueurs_onde_nm_arr, angles_rad_in_super_arr):
        if not longueurs_onde_nm_arr.size or not angles_rad_in_super_arr.size:
            return np.zeros((0, 0, 4))

        RT_results = np.zeros((len(longueurs_onde_nm_arr), len(angles_rad_in_super_arr), 4))

        for i_l, current_l_nm in enumerate(longueurs_onde_nm_arr):
            for i_a, current_theta_rad_super in enumerate(angles_rad_in_super_arr):
                alpha_snell = n_superstrate_real * np.sin(current_theta_rad_super)
                if abs(alpha_snell) > n_superstrate_real and not np.isclose(abs(alpha_snell), n_superstrate_real):
                    alpha_snell = np.sign(alpha_snell) * n_superstrate_real
                
                for pol_idx, pol_type in enumerate(['s', 'p']):
                    M_globale = np.eye(2, dtype=complex)
                    if emp_factors:
                        for i_couche, factor_qwt_couche in enumerate(emp_factors):
                            n_cplx_couche = nH if i_couche % 2 == 0 else nL
                            ep_phys_couche = ep_physical_nm[i_couche]
                            M_c = calcul_M_couche(pol_type, alpha_snell, n_cplx_couche, ep_phys_couche, current_l_nm)
                            M_globale = M_c @ M_globale 
                    
                    eta_super_sqrt_arg = n_superstrate_real**2 - alpha_snell**2
                    eta_super_sqrt = np.sqrt(eta_super_sqrt_arg + 0j)
                    eta_super_adm = (n_superstrate_real**2 / eta_super_sqrt if eta_super_sqrt != 0 else np.inf) if pol_type == 'p' else eta_super_sqrt
                    
                    eta_sub_sqrt_arg = nSub_complex**2 - alpha_snell**2
                    eta_sub_sqrt = np.sqrt(eta_sub_sqrt_arg + 0j)
                    eta_sub_adm = (nSub_complex**2 / eta_sub_sqrt if eta_sub_sqrt != 0 else np.inf) if pol_type == 'p' else eta_sub_sqrt

                    m00, m01 = M_globale[0,0], M_globale[0,1]
                    m10, m11 = M_globale[1,0], M_globale[1,1]

                    den_rt = (eta_super_adm * m00 + eta_sub_adm * m11 + eta_super_adm * eta_sub_adm * m01 + m10)
                    if den_rt == 0 or np.isinf(den_rt) or np.isnan(den_rt):
                        r_infini = np.inf
                        t_infini = 0
                    else:
                        r_infini = ((eta_super_adm * m00 - eta_sub_adm * m11 + eta_super_adm * eta_sub_adm * m01 - m10) / den_rt)
                        t_infini = (2 * eta_super_adm / den_rt)
                    
                    R_infini = np.abs(r_infini)**2 if not (np.isinf(r_infini) or np.isnan(r_infini)) else 1.0
                    
                    if np.real(eta_super_adm) != 0 and not (np.isinf(eta_super_adm) or np.isnan(eta_super_adm) or np.isinf(t_infini) or np.isnan(t_infini)):
                        T_infini = (np.real(eta_sub_adm) / np.real(eta_super_adm)) * np.abs(t_infini)**2
                    else:
                        T_infini = 0.0
                    
                    R_val, T_val = R_infini, T_infini

                    if substrat_fini:
                        den_Rb = (eta_sub_adm + eta_super_adm)
                        if den_Rb == 0 or np.isinf(den_Rb) or np.isnan(den_Rb):
                            Rb = 1.0
                        else:
                            Rb = np.abs((eta_sub_adm - eta_super_adm) / den_Rb)**2
                        
                        den_sub_fini = (1.0 - R_infini * Rb) 
                        if den_sub_fini != 0 and not (np.isinf(den_sub_fini) or np.isnan(den_sub_fini)):
                            R_val = R_infini + (T_infini**2 * Rb) / den_sub_fini
                            T_val = T_infini * (1.0 - Rb) / den_sub_fini
                        else:
                            R_val = R_infini
                            T_val = T_infini
                            
                    R_val = np.clip(R_val, 0, 1.001)
                    T_val = np.clip(T_val, 0, 1.001)

                    if pol_type == 's':
                        RT_results[i_l, i_a, 0] = R_val
                        RT_results[i_l, i_a, 2] = T_val
                    else: 
                        RT_results[i_l, i_a, 1] = R_val
                        RT_results[i_l, i_a, 3] = T_val
        return RT_results

    RT_spectral = calcul_RT_globale(l_nm, np.array([theta_inc_spectral_rad]))
    RT_angular = calcul_RT_globale(l_ang_nm, theta_inc_ang_rad)
    
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
