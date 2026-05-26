import streamlit as st
import pandas as pd
import os, sys, re
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(__file__))
from data_manager import (
    init_excel, load_sheet, save_sheet, get_config, save_config,
    get_sheets_names, EXCEL_FILE, SHEETS
)

st.set_page_config(
    page_title="💍 PCO — Hussen & Fransesca",
    page_icon="💍",
    layout="wide",
    initial_sidebar_state="expanded",
)

SAUGE      = "#87A878"
CHAMPAGNE  = "#C8A97E"
SAUGE_DARK = "#5E7D50"
CHAMP_DARK = "#A07840"
IVORY      = "#FDFAF4"
SAUGE_LIGHT= "#EBF2E8"
CHAMP_LIGHT= "#FAF3E8"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Lato:wght@300;400;700&display=swap');

html, body, [class*="css"] {{ font-family: 'Lato', sans-serif; background-color: {IVORY}; }}
h1,h2,h3 {{ font-family: 'Playfair Display', serif !important; }}
.main {{ background-color: {IVORY} !important; }}

.stSidebar > div:first-child {{
  background: linear-gradient(180deg, {SAUGE_DARK} 0%, #3d5c30 100%) !important;
}}
.stSidebar label, .stSidebar .stRadio label {{ color: white !important; }}

.hero-banner {{
  background: linear-gradient(135deg, {SAUGE_DARK} 0%, {CHAMPAGNE} 50%, {SAUGE_DARK} 100%);
  border-radius: 16px; padding: 28px 40px; text-align: center; color: white;
  margin-bottom: 24px; box-shadow: 0 8px 32px rgba(94,125,80,0.35);
}}
.hero-banner h1 {{ color: white !important; font-size: 2.2rem; margin: 0; letter-spacing: 2px; }}
.hero-banner p  {{ color: #FFF8EE; margin: 6px 0 0; font-size: 1.05rem; letter-spacing: 1px; }}

.section-title {{
  font-family: 'Playfair Display', serif;
  color: {SAUGE_DARK}; font-size: 1.5rem;
  border-bottom: 2px solid {CHAMPAGNE};
  padding-bottom: 8px; margin-bottom: 20px;
}}

.kpi-card {{
  background: white; border-left: 4px solid {CHAMPAGNE};
  border-radius: 10px; padding: 18px 20px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.07); text-align: center;
}}
.kpi-card .kpi-value {{ font-size: 2rem; font-weight: 700; color: {SAUGE_DARK}; font-family: 'Playfair Display', serif; }}
.kpi-card .kpi-label {{ font-size: 0.82rem; color: #888; margin-top: 4px; }}

.task-card {{
  background: white; border-radius: 10px;
  padding: 14px 18px; margin-bottom: 10px;
  border-left: 4px solid {CHAMPAGNE};
  box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}}
.task-card .t-name  {{ font-weight:700; color:{SAUGE_DARK}; font-size:1rem; }}
.task-card .t-role  {{ color:#666; font-size:0.85rem; }}
.task-card .t-task  {{ color:#333; font-size:0.9rem; margin-top:6px; }}
.task-card .t-phone {{ color:{CHAMP_DARK}; font-size:0.85rem; margin-top:4px; }}

.badge-done    {{ background:#E8F5E9; color:#2E7D32; padding:3px 10px; border-radius:12px; font-size:0.78rem; font-weight:600; }}
.badge-pending {{ background:#FFF8E6; color:#8B6000; padding:3px 10px; border-radius:12px; font-size:0.78rem; font-weight:600; }}
.badge-progress{{ background:#E3F2FD; color:#1565C0; padding:3px 10px; border-radius:12px; font-size:0.78rem; font-weight:600; }}

.stButton > button {{
  background: linear-gradient(135deg, {SAUGE_DARK}, {CHAMPAGNE}) !important;
  color: white !important; border: none !important;
  border-radius: 8px !important; font-weight: 600 !important;
  padding: 8px 20px !important;
}}
.stButton > button:hover {{ opacity:0.9; }}

.stTabs [data-baseweb="tab-list"] {{ gap: 8px; }}
.stTabs [data-baseweb="tab"] {{ border-radius: 6px 6px 0 0 !important; }}
.stTabs [aria-selected="true"] {{ background: {SAUGE_LIGHT} !important; color: {SAUGE_DARK} !important; font-weight:700; }}

div[data-testid="stMetricValue"] {{ color: {SAUGE_DARK} !important; }}

/* Boutons icône petits */
.btn-sm > button {{
  padding: 2px 8px !important;
  font-size: 0.75rem !important;
  min-height: 28px !important;
}}
</style>
""", unsafe_allow_html=True)

# ── Init ─────────────────────────────────────────────────────────────────────
init_excel()
wedding_date = date(2025, 8, 1)
today = date.today()

# ── Helpers ──────────────────────────────────────────────────────────────────
def hero():
    st.markdown("""
    <div class="hero-banner">
      <h1>💍 Hussen Junior &amp; Fransesca 💍</h1>
      <p>✨ Gestion PCO — Mariage du 1er Août 2025 ✨</p>
    </div>""", unsafe_allow_html=True)

def section(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)

def badge(status):
    s = str(status) if status else ""
    if re.search(r"fait|confirmé|oui|✅|actif|terminé|signé", s, re.I):
        return "done", f'<span class="badge-done">✅ {s}</span>'
    elif re.search(r"cours|progress", s, re.I):
        return "progress", f'<span class="badge-progress">🔄 {s}</span>'
    else:
        return "pending", f'<span class="badge-pending">⏳ {s}</span>'

def safe(val):
    v = str(val) if val is not None else ""
    return "" if v in ("nan", "NaN", "None") else v

def force_str_cols(df, cols):
    """Force les colonnes textuelles en object pour éviter l'erreur float64."""
    for col in cols:
        if col in df.columns:
            df[col] = df[col].astype(object)
    return df

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center; padding:10px 0 20px;'>
      <div style='font-size:3rem;'>💍</div>
      <div style='color:white; font-size:1.1rem; font-weight:700; letter-spacing:1px;'>PCO Mariage</div>
      <div style='color:#EEE8DC; font-size:0.85rem;'>Hussen &amp; Fransesca</div>
      <div style='color:#EEE8DC; font-size:0.8rem; margin-top:4px;'>📅 1er Août 2025</div>
    </div>""", unsafe_allow_html=True)

    page = st.radio("Nav", [
        "🏠 Tableau de bord",
        "👥 Équipe & Tâches",
        "📋 Déroulement",
        "✅ Checklist",
        "💰 Budget",
        "🎉 Invités",
        "🤝 Fournisseurs",
        "⚙️ Configuration",
        "📝 Notes PCO",
    ], label_visibility="collapsed")

    st.markdown("---")
    delta = (wedding_date - today).days
    if delta > 0:
        st.markdown(f"""<div style='text-align:center; color:white;'>
          <div style='font-size:2.5rem; font-weight:700; color:#FFD700;'>{delta}</div>
          <div style='font-size:0.85rem; color:#EEE8DC;'>jours avant le grand jour</div>
        </div>""", unsafe_allow_html=True)
    elif delta == 0:
        st.markdown("<div style='text-align:center;color:#FFD700;font-size:1.2rem;'>🎉 C'est aujourd'hui !</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div style='text-align:center;color:#EEE8DC;font-size:0.9rem;'>✨ Mariage passé — Félicitations !</div>", unsafe_allow_html=True)

    st.markdown("---")
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("📥 Télécharger Excel", data=f,
                file_name=EXCEL_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: TABLEAU DE BORD
# ═══════════════════════════════════════════════════════════════════════════
if page == "🏠 Tableau de bord":
    hero()
    config = get_config()
    df_check  = load_sheet(SHEETS["Checklist"])
    df_equipe = load_sheet(SHEETS["Equipe"])
    df_invites= load_sheet(SHEETS["Invites"])

    total_items, done_items, done_col = 0, 0, None
    if not df_check.empty:
        total_items = len(df_check)
        for col in df_check.columns:
            if "statut" in str(col).lower():
                done_col = col; break
        if done_col:
            done_items = df_check[done_col].astype(str).str.contains("fait|confirmé|oui|✅", case=False, na=False).sum()

    equipe_count  = len(df_equipe.dropna(how='all'))  if not df_equipe.empty  else 0
    invites_count = len(df_invites.dropna(how='all')) if not df_invites.empty else 0
    pct_global    = int(done_items / total_items * 100) if total_items > 0 else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    for col, val, lbl in [
        (c1, max(delta,0), "📅 Jours restants"),
        (c2, f"{pct_global}%", "✅ Checklist"),
        (c3, equipe_count, "👥 Équipe"),
        (c4, invites_count, "🎉 Invités"),
        (c5, "🎨", config.get("Thème couleurs","Sauge & Champagne")),
    ]:
        col.markdown(f'<div class="kpi-card"><div class="kpi-value">{val}</div><div class="kpi-label">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        section("🎨 Palette de couleurs")
        swatches = [
            (SAUGE_DARK, "Vert Sauge", SAUGE_DARK),
            (CHAMPAGNE,  "Champagne", CHAMP_DARK),
            (SAUGE,      "Sauge clair", SAUGE_DARK),
            ("#FDFAF4",  "Ivoire", "#999"),
        ]
        cols_sw = st.columns(4)
        for i,(bg,name,txt) in enumerate(swatches):
            cols_sw[i].markdown(f"""<div style='text-align:center;'>
              <div style='background:{bg};width:56px;height:56px;border-radius:50%;margin:auto;border:2px solid #ccc;'></div>
              <div style='font-size:0.75rem;margin-top:5px;color:{txt};font-weight:600;'>{name}</div>
              <div style='font-size:0.7rem;color:#aaa;'>{bg}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        section("📋 Infos clés")
        for lbl, key, defval in [
            ("💒 Marié","Marié","Hussen Junior"),("👰 Mariée","Mariée","Fransesca"),
            ("📅 Date","Date du mariage","01/08/2025"),
            ("📍 Cérémonie","Lieu cérémonie","À définir"),
            ("🥂 Réception","Lieu réception","À définir"),
            ("💰 Budget","Budget total estimé","À définir"),
        ]:
            st.markdown(f"**{lbl}** : {config.get(key, defval) or defval}")

    with col2:
        section("✅ Avancement par catégorie")
        if not df_check.empty and "Catégorie" in df_check.columns and done_col:
            for cat, grp in df_check.groupby("Catégorie"):
                t = len(grp)
                d = grp[done_col].astype(str).str.contains("fait|confirmé|oui|✅", case=False, na=False).sum()
                p = int(d/t*100) if t>0 else 0
                bar_col = "#4CAF50" if p==100 else CHAMPAGNE if p>0 else "#E0E0E0"
                st.markdown(f"""<div style='display:flex;align-items:center;gap:8px;margin-bottom:5px;'>
                  <div style='width:130px;font-size:0.78rem;color:#555;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>{cat[:22]}</div>
                  <div style='flex:1;background:#F0F0F0;border-radius:4px;height:13px;'>
                    <div style='width:{p}%;background:{bar_col};height:13px;border-radius:4px;'></div>
                  </div>
                  <div style='width:38px;font-size:0.78rem;font-weight:700;color:{SAUGE_DARK};'>{p}%</div>
                </div>""", unsafe_allow_html=True)
        else:
            st.info("Remplis la checklist pour voir l'avancement.")

    # ── Programme Jour J ──────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    section("📅 Programme Jour J — Vue rapide")
    df_prog = load_sheet(SHEETS["Deroulement"])

    if df_prog.empty:
        st.info("Aucune étape dans le déroulement. Rendez-vous dans la page 📋 Déroulement pour en ajouter.")
    else:
        # Contrôles de gestion rapide
        with st.expander("⚙️ Gérer les étapes du programme", expanded=False):
            st.caption("Sélectionne une étape pour la modifier, supprimer ou déplacer.")
            etapes_labels = [f"{safe(r.iloc[2])} — {safe(r.iloc[0])}" for _, r in df_prog.iterrows()]
            sel_prog = st.selectbox("Étape à gérer", etapes_labels, key="prog_select")
            idx_prog = etapes_labels.index(sel_prog)

            act_col1, act_col2, act_col3, act_col4 = st.columns(4)

            # Monter
            if act_col1.button("⬆️ Monter", key="prog_up"):
                if idx_prog > 0:
                    df_prog.iloc[[idx_prog-1, idx_prog]] = df_prog.iloc[[idx_prog, idx_prog-1]].values
                    df_prog = df_prog.reset_index(drop=True)
                    save_sheet(SHEETS["Deroulement"], df_prog)
                    st.success("✅ Étape déplacée vers le haut."); st.rerun()
                else:
                    st.warning("Déjà en première position.")

            # Descendre
            if act_col2.button("⬇️ Descendre", key="prog_down"):
                if idx_prog < len(df_prog) - 1:
                    df_prog.iloc[[idx_prog, idx_prog+1]] = df_prog.iloc[[idx_prog+1, idx_prog]].values
                    df_prog = df_prog.reset_index(drop=True)
                    save_sheet(SHEETS["Deroulement"], df_prog)
                    st.success("✅ Étape déplacée vers le bas."); st.rerun()
                else:
                    st.warning("Déjà en dernière position.")

            # Supprimer
            if act_col3.button("🗑️ Supprimer", key="prog_del"):
                st.session_state["confirm_prog_del"] = idx_prog

            if st.session_state.get("confirm_prog_del") == idx_prog:
                st.warning(f"⚠️ Supprimer **{sel_prog}** ?")
                cc1, cc2 = st.columns(2)
                if cc1.button("✅ Oui, supprimer", key="prog_del_ok"):
                    df_prog = df_prog.drop(index=idx_prog).reset_index(drop=True)
                    save_sheet(SHEETS["Deroulement"], df_prog)
                    st.session_state.pop("confirm_prog_del", None)
                    st.success("🗑️ Étape supprimée."); st.rerun()
                if cc2.button("❌ Annuler", key="prog_del_cancel"):
                    st.session_state.pop("confirm_prog_del", None)
                    st.rerun()

            # Modifier directement
            act_col4.write("")
            st.markdown("**✏️ Modifier cette étape :**")
            row_prog = df_prog.iloc[idx_prog]
            with st.form("edit_prog_quick"):
                ep1, ep2 = st.columns(2)
                new_etape = ep1.text_input("Étape", value=safe(row_prog.iloc[0]))
                new_heure = ep2.text_input("Heure", value=safe(row_prog.iloc[2]))
                ep3, ep4 = st.columns(2)
                new_lieu  = ep3.text_input("Lieu", value=safe(row_prog.iloc[3]))
                new_resp  = ep4.text_input("Responsable", value=safe(row_prog.iloc[4]))
                new_desc  = st.text_area("Description", value=safe(row_prog.iloc[1]), height=60)
                opts_s = ["À planifier","Confirmé","En cours","Terminé"]
                cur_s  = safe(row_prog.iloc[7]) if len(row_prog)>7 else "À planifier"
                ep5, ep6 = st.columns(2)
                new_st   = ep5.selectbox("Statut", opts_s, index=opts_s.index(cur_s) if cur_s in opts_s else 0)
                new_tenue= ep6.text_input("Tenue", value=safe(row_prog.iloc[5]) if len(row_prog)>5 else "")
                if st.form_submit_button("💾 Enregistrer les modifications"):
                    df_prog = force_str_cols(df_prog, list(df_prog.columns))
                    df_prog.iat[idx_prog, 0] = new_etape
                    df_prog.iat[idx_prog, 1] = new_desc
                    df_prog.iat[idx_prog, 2] = new_heure
                    df_prog.iat[idx_prog, 3] = new_lieu
                    df_prog.iat[idx_prog, 4] = new_resp
                    if len(df_prog.columns) > 5: df_prog.iat[idx_prog, 5] = new_tenue
                    if len(df_prog.columns) > 7: df_prog.iat[idx_prog, 7] = new_st
                    save_sheet(SHEETS["Deroulement"], df_prog)
                    st.success("✅ Étape mise à jour !"); st.rerun()

        # Affichage des blocs
        ca, cb = st.columns(2)
        half = len(df_prog)//2
        for i,(_, row) in enumerate(df_prog.iterrows()):
            tgt = ca if i < half else cb
            with tgt:
                etape = safe(row.iloc[0]); heure = safe(row.iloc[2]); lieu = safe(row.iloc[3])
                st.markdown(f"""<div style='display:flex;gap:12px;background:white;border-radius:10px;
                  padding:10px 14px;margin-bottom:8px;box-shadow:0 1px 6px rgba(0,0,0,0.06);
                  border-left:4px solid {SAUGE_DARK};'>
                  <div style='font-weight:700;color:{CHAMPAGNE};min-width:60px;font-size:0.95rem;'>{heure}</div>
                  <div>
                    <div style='font-weight:600;color:{SAUGE_DARK};'>{etape}</div>
                    <div style='font-size:0.8rem;color:#777;'>📍 {lieu}</div>
                  </div>
                </div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: ÉQUIPE & TÂCHES
# ═══════════════════════════════════════════════════════════════════════════
elif page == "👥 Équipe & Tâches":
    hero()
    section("👥 Équipe du Mariage & Tâches Assignées")

    df = load_sheet(SHEETS["Equipe"])
    COLS_EQUIPE = ["Nom", "Rôle / Fonction", "Tâches Assignées", "Téléphone", "Email", "Statut", "Notes"]
    if df.empty or list(df.columns) != COLS_EQUIPE:
        if not df.empty:
            df.columns = COLS_EQUIPE[:len(df.columns)]
        else:
            df = pd.DataFrame(columns=COLS_EQUIPE)

    tab_vue, tab_edit = st.tabs(["📋 Vue Équipe", "✏️ Modifier / Ajouter"])

    with tab_vue:
        rows_clean = [(i, row) for i, row in df.iterrows()]
        if rows_clean:
            cols2 = st.columns(2)
            for idx_card, (df_idx, row) in enumerate(rows_clean):
                nom    = safe(row.get("Nom",             row.iloc[0] if len(row)>0 else ""))
                role   = safe(row.get("Rôle / Fonction", row.iloc[1] if len(row)>1 else ""))
                tache  = safe(row.get("Tâches Assignées",row.iloc[2] if len(row)>2 else ""))
                tel    = safe(row.get("Téléphone",       row.iloc[3] if len(row)>3 else ""))
                email  = safe(row.get("Email",           row.iloc[4] if len(row)>4 else ""))
                statut = safe(row.get("Statut",          row.iloc[5] if len(row)>5 else ""))
                notes  = safe(row.get("Notes",           row.iloc[6] if len(row)>6 else ""))
                nom_disp = nom if nom else "— À renseigner —"
                _, badge_html = badge(statut)
                notes_html = f'<div style="color:#999;font-size:0.8rem;margin-top:4px;font-style:italic;">{notes}</div>' if notes else ""
                with cols2[idx_card % 2]:
                    st.markdown(f"""
                    <div class="task-card">
                      <div class="t-name">👤 {nom_disp}</div>
                      <div class="t-role">🎯 {role}</div>
                      <div class="t-task">📌 {tache if tache else '—'}</div>
                      <div class="t-phone">📞 {tel if tel else '—'} &nbsp;|&nbsp; ✉️ {email if email else '—'}</div>
                      <div style="margin-top:6px;">{badge_html}</div>
                      {notes_html}
                    </div>""", unsafe_allow_html=True)
        else:
            st.info("Aucun membre d'équipe. Utilisez l'onglet Modifier / Ajouter.")

    with tab_edit:
        st.markdown(f"### ➕ Ajouter un nouveau membre")
        with st.form("form_add_team", clear_on_submit=True):
            a1, a2 = st.columns(2)
            new_nom    = a1.text_input("Nom & Prénom *")
            new_role   = a2.text_input("Rôle / Fonction *")
            new_tache  = st.text_area("Tâches assignées", height=70)
            b1, b2 = st.columns(2)
            new_tel    = b1.text_input("Téléphone")
            new_email  = b2.text_input("Email")
            c1_, c2_ = st.columns(2)
            new_statut = c1_.selectbox("Statut", ["Actif", "À confirmer", "Indisponible"])
            new_notes  = c2_.text_input("Notes")
            if st.form_submit_button("💾 Ajouter au fichier Excel"):
                if new_nom.strip() and new_role.strip():
                    row_add = pd.DataFrame([[new_nom, new_role, new_tache, new_tel, new_email, new_statut, new_notes]],
                                            columns=COLS_EQUIPE)
                    df = pd.concat([df, row_add], ignore_index=True)
                    save_sheet(SHEETS["Equipe"], df)
                    st.success(f"✅ {new_nom} ajouté(e) !")
                    st.rerun()
                else:
                    st.error("Nom et Rôle sont obligatoires.")

        st.markdown("---")

        if df.empty:
            st.info("Aucun membre à modifier.")
        else:
            st.markdown("### ✏️ Modifier un membre existant")
            labels = []
            for i, row in df.iterrows():
                n = safe(row.get("Nom", row.iloc[0] if len(row)>0 else ""))
                r = safe(row.get("Rôle / Fonction", row.iloc[1] if len(row)>1 else ""))
                label = f"{r}" if not n else f"{n} — {r}"
                labels.append((i, label))

            label_list = [l for _, l in labels]
            sel_label  = st.selectbox("Choisir le membre à modifier", label_list)
            sel_df_idx = next(i for i, l in labels if l == sel_label)
            row_sel    = df.loc[sel_df_idx]

            with st.form("form_edit_team"):
                e1, e2 = st.columns(2)
                e_nom   = e1.text_input("Nom & Prénom",
                    value=safe(row_sel.get("Nom", row_sel.iloc[0] if len(row_sel)>0 else "")))
                e_role  = e2.text_input("Rôle / Fonction",
                    value=safe(row_sel.get("Rôle / Fonction", row_sel.iloc[1] if len(row_sel)>1 else "")))
                e_tache = st.text_area("Tâches assignées", height=70,
                    value=safe(row_sel.get("Tâches Assignées", row_sel.iloc[2] if len(row_sel)>2 else "")))
                f1, f2 = st.columns(2)
                e_tel   = f1.text_input("Téléphone",
                    value=safe(row_sel.get("Téléphone", row_sel.iloc[3] if len(row_sel)>3 else "")))
                e_email = f2.text_input("Email",
                    value=safe(row_sel.get("Email", row_sel.iloc[4] if len(row_sel)>4 else "")))
                g1, g2 = st.columns(2)
                statuts_list = ["Actif", "À confirmer", "Indisponible"]
                cur_s = safe(row_sel.get("Statut", row_sel.iloc[5] if len(row_sel)>5 else ""))
                s_idx = statuts_list.index(cur_s) if cur_s in statuts_list else 1
                e_statut = g1.selectbox("Statut", statuts_list, index=s_idx)
                e_notes  = g2.text_input("Notes",
                    value=safe(row_sel.get("Notes", row_sel.iloc[6] if len(row_sel)>6 else "")))
                h1, h2 = st.columns(2)
                btn_save = h1.form_submit_button("💾 Enregistrer les modifications")
                btn_del  = h2.form_submit_button("🗑️ Supprimer ce membre")

                if btn_save:
                    # ── FIX float64 : forcer toutes les colonnes en object ──
                    df = force_str_cols(df, COLS_EQUIPE)
                    df.at[sel_df_idx, "Nom"]              = e_nom
                    df.at[sel_df_idx, "Rôle / Fonction"]  = e_role
                    df.at[sel_df_idx, "Tâches Assignées"] = e_tache
                    df.at[sel_df_idx, "Téléphone"]        = e_tel
                    df.at[sel_df_idx, "Email"]            = e_email
                    df.at[sel_df_idx, "Statut"]           = e_statut
                    df.at[sel_df_idx, "Notes"]            = e_notes
                    save_sheet(SHEETS["Equipe"], df)
                    st.success("✅ Informations enregistrées !")
                    st.rerun()
                if btn_del:
                    df = df.drop(index=sel_df_idx).reset_index(drop=True)
                    save_sheet(SHEETS["Equipe"], df)
                    st.success("🗑️ Membre supprimé.")
                    st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: DÉROULEMENT — VERSION COMPLÈTE AVEC ÉDITION / RÉORG / SUPPRESSION
# ═══════════════════════════════════════════════════════════════════════════
elif page == "📋 Déroulement":
    hero()
    section("📋 Déroulement Complet du Mariage")
    df = load_sheet(SHEETS["Deroulement"])

    # S'assurer que le df a assez de colonnes nommées
    COLS_DERO = ["Étape","Description","Heure","Lieu","Responsable","Tenue","Photos","Statut","Notes"]
    if not df.empty and len(df.columns) >= len(COLS_DERO):
        df.columns = list(COLS_DERO) + list(df.columns[len(COLS_DERO):])
    elif not df.empty:
        # Compléter avec des colonnes vides si nécessaire
        for i in range(len(df.columns), len(COLS_DERO)):
            df[COLS_DERO[i]] = ""
        df.columns = COLS_DERO

    tab1, tab2, tab3 = st.tabs(["🕐 Timeline & Gestion", "📝 Tableau complet", "➕ Ajouter une étape"])

    # ─── ONGLET 1 : Timeline interactive avec gestion complète ───────────────
    with tab1:
        if df.empty:
            st.info("Aucune étape planifiée. Utilisez l'onglet ➕ Ajouter une étape.")
        else:
            st.markdown("""
            <div style='background:#EBF2E8;border-radius:8px;padding:10px 16px;margin-bottom:16px;font-size:0.85rem;color:#5E7D50;'>
            💡 <b>Chaque étape</b> possède des boutons pour la <b>modifier</b>, la <b>déplacer</b> ou la <b>supprimer</b>.
            Développe le panneau ✏️ d'une étape pour éditer tous ses champs directement.
            </div>
            """, unsafe_allow_html=True)

            for idx, (df_idx, row) in enumerate(df.iterrows()):
                etape  = safe(row["Étape"])   if "Étape"       in df.columns else safe(row.iloc[0])
                desc   = safe(row["Description"]) if "Description" in df.columns else safe(row.iloc[1])
                heure  = safe(row["Heure"])   if "Heure"       in df.columns else safe(row.iloc[2])
                lieu   = safe(row["Lieu"])    if "Lieu"        in df.columns else safe(row.iloc[3])
                resp   = safe(row["Responsable"]) if "Responsable" in df.columns else safe(row.iloc[4])
                tenue  = safe(row["Tenue"])   if "Tenue"       in df.columns else (safe(row.iloc[5]) if len(row)>5 else "")
                photos = safe(row["Photos"])  if "Photos"      in df.columns else (safe(row.iloc[6]) if len(row)>6 else "")
                statut = safe(row["Statut"])  if "Statut"      in df.columns else (safe(row.iloc[7]) if len(row)>7 else "")
                notes  = safe(row["Notes"])   if "Notes"       in df.columns else (safe(row.iloc[8]) if len(row)>8 else "")

                _, bdg = badge(statut)

                # ── Carte avec barre d'actions ──
                sub = " &nbsp;·&nbsp; ".join(filter(None, [
                    f"📍 {lieu}"    if lieu  else "",
                    f"👤 {resp}"    if resp  else "",
                    f"👗 {tenue}"   if tenue else "",
                    f"📸 Photos"    if photos and photos.lower() not in ("non","") else "",
                ]))

                st.markdown(f"""
                <div style='display:flex;gap:16px;margin-bottom:4px;background:white;border-radius:12px;
                  padding:13px 18px;box-shadow:0 2px 10px rgba(0,0,0,0.06);border-left:4px solid {SAUGE_DARK};'>
                  <div style='min-width:75px;'>
                    <div style='font-weight:700;color:{CHAMPAGNE};font-size:1.1rem;'>{heure or "—"}</div>
                    <div style='margin-top:5px;'>{bdg}</div>
                  </div>
                  <div style='flex:1;'>
                    <div style='font-weight:700;color:{SAUGE_DARK};font-size:0.98rem;'>{etape or "Sans titre"}</div>
                    <div style='color:#555;font-size:0.85rem;margin-top:3px;'>{desc}</div>
                    <div style='font-size:0.8rem;color:#777;margin-top:6px;'>{sub}</div>
                    {f'<div style="font-size:0.78rem;color:#aaa;font-style:italic;margin-top:3px;">📝 {notes}</div>' if notes else ''}
                  </div>
                </div>""", unsafe_allow_html=True)

                # ── Barre d'actions ──
                btn_cols = st.columns([1,1,1,1,4])
                key_prefix = f"dero_{idx}"

                # Monter
                if btn_cols[0].button("⬆️", key=f"{key_prefix}_up", help="Monter"):
                    if idx > 0:
                        idxs = list(df.index)
                        idxs[idx-1], idxs[idx] = idxs[idx], idxs[idx-1]
                        df = df.loc[idxs].reset_index(drop=True)
                        save_sheet(SHEETS["Deroulement"], df)
                        st.rerun()

                # Descendre
                if btn_cols[1].button("⬇️", key=f"{key_prefix}_down", help="Descendre"):
                    if idx < len(df)-1:
                        idxs = list(df.index)
                        idxs[idx], idxs[idx+1] = idxs[idx+1], idxs[idx]
                        df = df.loc[idxs].reset_index(drop=True)
                        save_sheet(SHEETS["Deroulement"], df)
                        st.rerun()

                # Supprimer
                if btn_cols[2].button("🗑️", key=f"{key_prefix}_del", help="Supprimer"):
                    st.session_state[f"del_confirm_{idx}"] = True

                if st.session_state.get(f"del_confirm_{idx}"):
                    st.warning(f"⚠️ Supprimer **{etape}** ? Cette action est irréversible.")
                    dc1, dc2 = st.columns(2)
                    if dc1.button("✅ Confirmer la suppression", key=f"{key_prefix}_del_yes"):
                        df = df.drop(index=df_idx).reset_index(drop=True)
                        save_sheet(SHEETS["Deroulement"], df)
                        st.session_state.pop(f"del_confirm_{idx}", None)
                        st.rerun()
                    if dc2.button("❌ Annuler", key=f"{key_prefix}_del_no"):
                        st.session_state.pop(f"del_confirm_{idx}", None)
                        st.rerun()

                # Modifier (expander)
                with btn_cols[3].container():
                    pass  # espace visuel

                # Expander d'édition
                with st.expander(f"✏️ Modifier : {etape or 'cette étape'}"):
                    with st.form(f"edit_dero_{idx}_{df_idx}"):
                        ed1, ed2, ed3 = st.columns(3)
                        new_etape = ed1.text_input("Étape *", value=etape)
                        new_heure = ed2.text_input("Heure (ex: 10:30)", value=heure)
                        new_lieu  = ed3.text_input("Lieu", value=lieu)

                        new_desc  = st.text_area("Description", value=desc, height=70)

                        ed4, ed5, ed6 = st.columns(3)
                        new_resp  = ed4.text_input("Responsable", value=resp)
                        new_tenue = ed5.text_input("Tenue", value=tenue)
                        new_photos = ed6.selectbox("Photos", ["Non","Oui"],
                                        index=1 if photos and photos.lower()=="oui" else 0)

                        opts_st = ["À planifier","Confirmé","En cours","Terminé"]
                        cur_idx = opts_st.index(statut) if statut in opts_st else 0
                        ed7, ed8 = st.columns(2)
                        new_statut = ed7.selectbox("Statut", opts_st, index=cur_idx)
                        new_notes  = ed8.text_input("Notes", value=notes)

                        if st.form_submit_button("💾 Enregistrer"):
                            if new_etape.strip():
                                df = force_str_cols(df, list(df.columns))
                                df.at[df_idx, "Étape"]       = new_etape
                                df.at[df_idx, "Description"] = new_desc
                                df.at[df_idx, "Heure"]       = new_heure
                                df.at[df_idx, "Lieu"]        = new_lieu
                                df.at[df_idx, "Responsable"] = new_resp
                                df.at[df_idx, "Tenue"]       = new_tenue
                                df.at[df_idx, "Photos"]      = new_photos
                                df.at[df_idx, "Statut"]      = new_statut
                                df.at[df_idx, "Notes"]       = new_notes
                                save_sheet(SHEETS["Deroulement"], df)
                                st.success(f"✅ '{new_etape}' mis à jour !")
                                st.rerun()
                            else:
                                st.error("Le nom de l'étape est obligatoire.")

                st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    # ─── ONGLET 2 : Tableau complet ──────────────────────────────────────────
    with tab2:
        if not df.empty:
            st.dataframe(df.fillna(""), use_container_width=True, hide_index=True)
        else:
            st.info("Aucune étape.")

    # ─── ONGLET 3 : Ajouter une étape ────────────────────────────────────────
    with tab3:
        st.markdown("### ➕ Ajouter une nouvelle étape")
        with st.form("add_step"):
            r1c1, r1c2, r1c3 = st.columns(3)
            e_n = r1c1.text_input("Étape *")
            h_n = r1c2.text_input("Heure", placeholder="10:30")
            l_n = r1c3.text_input("Lieu")
            d_n = st.text_area("Description", height=60)
            r2c1, r2c2, r2c3 = st.columns(3)
            re_n = r2c1.text_input("Responsable")
            te_n = r2c2.text_input("Tenue")
            ph_n = r2c3.selectbox("Photos", ["Non","Oui"])
            r3c1, r3c2 = st.columns(2)
            st_n = r3c1.selectbox("Statut", ["À planifier","Confirmé","En cours","Terminé"])
            no_n = r3c2.text_input("Notes")

            pos_opts = ["À la fin"] + ([f"Avant : {safe(r['Étape'])}" for _, r in df.iterrows() if safe(r["Étape"])] if not df.empty else [])
            pos_sel  = st.selectbox("📌 Position", pos_opts)

            if st.form_submit_button("💾 Ajouter l'étape"):
                if e_n.strip():
                    new_row = pd.DataFrame([[e_n, d_n, h_n, l_n, re_n, te_n, ph_n, st_n, no_n]],
                                            columns=COLS_DERO)
                    if pos_sel == "À la fin" or df.empty:
                        df = pd.concat([df, new_row], ignore_index=True)
                    else:
                        # Insérer avant l'étape choisie
                        ref_etape = pos_sel.replace("Avant : ", "")
                        ref_idx = next((i for i, r in df.iterrows() if safe(r["Étape"]) == ref_etape), None)
                        if ref_idx is not None:
                            top    = df.iloc[:ref_idx]
                            bottom = df.iloc[ref_idx:]
                            df = pd.concat([top, new_row, bottom], ignore_index=True)
                        else:
                            df = pd.concat([df, new_row], ignore_index=True)
                    save_sheet(SHEETS["Deroulement"], df)
                    st.success(f"✅ Étape '{e_n}' ajoutée !")
                    st.rerun()
                else:
                    st.error("Le nom de l'étape est obligatoire.")


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: CHECKLIST
# ═══════════════════════════════════════════════════════════════════════════
elif page == "✅ Checklist":
    hero()
    section("✅ Checklist Complète du Mariage")
    df = load_sheet(SHEETS["Checklist"])

    if df.empty:
        st.warning("Aucune donnée trouvée."); st.stop()

    statut_col = next((c for c in df.columns if "statut" in str(c).lower()), None)
    if statut_col:
        total = len(df)
        done  = df[statut_col].astype(str).str.contains("fait|confirmé|oui|✅", case=False, na=False).sum()
        pct   = int(done/total*100) if total>0 else 0
        bar_col = "#4CAF50" if pct==100 else CHAMPAGNE if pct>0 else "#E0E0E0"
        st.markdown(f"""
        <div style='background:white;padding:16px 20px;border-radius:12px;margin-bottom:20px;box-shadow:0 2px 10px rgba(0,0,0,0.07);'>
          <div style='display:flex;justify-content:space-between;margin-bottom:10px;'>
            <span style='font-weight:600;color:{SAUGE_DARK};font-size:1.05rem;'>Progression globale</span>
            <span style='font-weight:700;color:{CHAMP_DARK};font-size:1.2rem;'>{done}/{total} — {pct}%</span>
          </div>
          <div style='background:#F0F0F0;border-radius:8px;height:18px;'>
            <div style='width:{pct}%;background:linear-gradient(90deg,{SAUGE_DARK},{CHAMPAGNE});height:18px;border-radius:8px;'></div>
          </div>
        </div>""", unsafe_allow_html=True)

    cats = ["Toutes"] + sorted(df["Catégorie"].dropna().unique().tolist()) if "Catégorie" in df.columns else ["Toutes"]
    cat_filter = st.selectbox("Filtrer par catégorie", cats)
    show_df = df if cat_filter == "Toutes" else df[df["Catégorie"] == cat_filter]

    tab_cards, tab_table = st.tabs(["🃏 Vue cartes", "📊 Tableau"])

    with tab_cards:
        for _, row in show_df.iterrows():
            cat    = safe(row.iloc[0]); elem   = safe(row.iloc[1]); detail = safe(row.iloc[2])
            lieu_f = safe(row.iloc[3]); contact= safe(row.iloc[4]); prix   = safe(row.iloc[5])
            statut = safe(row.iloc[6]) if len(row)>6 else "À confirmer"
            is_done = bool(re.search(r"fait|confirmé|oui|✅", statut, re.I))
            border = "#4CAF50" if is_done else "#E0E0E0"
            badge_type, _ = badge(statut)
            if badge_type == "done":
                bdg_html = f'<span class="badge-done">✅ {statut}</span>'
            elif badge_type == "progress":
                bdg_html = f'<span class="badge-progress">🔄 {statut}</span>'
            else:
                bdg_html = f'<span class="badge-pending">⏳ {statut}</span>'

            sub_parts = []
            if lieu_f:    sub_parts.append(f"📍 {lieu_f}")
            if contact:   sub_parts.append(f"📞 {contact}")
            if prix and prix not in ("0",""):
                sub_parts.append(f"💰 {prix} FCFA")
            sub_line = " &nbsp;·&nbsp; ".join(sub_parts)

            st.markdown(f"""
            <div style='background:white;border-left:5px solid {border};border-radius:10px;
              padding:12px 16px;margin-bottom:8px;box-shadow:0 1px 6px rgba(0,0,0,0.06);
              display:flex;justify-content:space-between;align-items:center;'>
              <div style='flex:1;'>
                <div style='font-weight:600;color:#333;font-size:0.93rem;'>{cat} — {elem}</div>
                {'<div style="font-size:0.82rem;color:#666;margin-top:2px;">'+detail+'</div>' if detail else ''}
                {'<div style="font-size:0.78rem;color:#999;margin-top:4px;">'+sub_line+'</div>' if sub_line else ''}
              </div>
              <div style='margin-left:14px;white-space:nowrap;'>{bdg_html}</div>
            </div>""", unsafe_allow_html=True)

    with tab_table:
        st.dataframe(show_df.fillna(""), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### ✏️ Mettre à jour un élément")
    if len(df.columns) > 1:
        elem_labels = [f"{safe(r.iloc[0])} — {safe(r.iloc[1])}" for _, r in df.iterrows()]
        sel_e = st.selectbox("Élément à mettre à jour", elem_labels)
        sel_i = elem_labels.index(sel_e)
        row_s = df.iloc[sel_i]
        with st.form("upd_checklist"):
            u1,u2,u3 = st.columns(3)
            opts_s = ["À confirmer","En cours","Confirmé ✅","Fait ✅"]
            cur_st = safe(row_s.iloc[6]) if len(row_s)>6 else "À confirmer"
            new_st = u1.selectbox("Statut", opts_s, index=opts_s.index(cur_st) if cur_st in opts_s else 0)
            new_li = u2.text_input("Lieu / Fournisseur", value=safe(row_s.iloc[3]) if len(row_s)>3 else "")
            new_co = u3.text_input("Contact", value=safe(row_s.iloc[4]) if len(row_s)>4 else "")
            v1,v2 = st.columns(2)
            new_px = v1.text_input("Prix estimé (FCFA)", value=safe(row_s.iloc[5]) if len(row_s)>5 else "")
            new_no = v2.text_input("Notes", value=safe(row_s.iloc[8]) if len(row_s)>8 else "")
            if st.form_submit_button("💾 Valider"):
                df = force_str_cols(df, list(df.columns))
                for ci, nv in [(3,new_li),(4,new_co),(5,new_px),(6,new_st),(8,new_no)]:
                    if ci < len(df.columns): df.iat[sel_i, ci] = nv
                save_sheet(SHEETS["Checklist"], df)
                st.success(f"✅ Mis à jour : {sel_e}"); st.rerun()

    st.markdown("### ➕ Ajouter un élément")
    with st.form("add_check"):
        w1,w2 = st.columns(2)
        new_cat = w1.selectbox("Catégorie", [
            "📍 Lieu","🍽️ Traiteur","🥂 Boissons","🎂 Gâteau","🌸 Décoration",
            "👗 Tenues","💄 Beauté","📸 Photo/Vidéo","🎵 Musique",
            "💌 Papeterie","🖼️ Affiches","🚗 Transport","🏨 Hébergement",
            "🎁 Cadeaux","📄 Administratif","🎆 Divers","Autre"])
        new_el = w2.text_input("Élément *")
        x1,x2 = st.columns(2)
        new_det = x1.text_input("Détails"); new_st2 = x2.selectbox("Statut initial", ["À confirmer","En cours"])
        if st.form_submit_button("➕ Ajouter"):
            if new_el:
                nr = pd.DataFrame([[new_cat,new_el,new_det,"","","",new_st2,"",""]],
                                   columns=df.columns[:9] if len(df.columns)>=9 else list(range(9)))
                df = pd.concat([df,nr],ignore_index=True)
                save_sheet(SHEETS["Checklist"],df)
                st.success(f"✅ '{new_el}' ajouté !"); st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: BUDGET
# ═══════════════════════════════════════════════════════════════════════════
elif page == "💰 Budget":
    hero()
    section("💰 Suivi du Budget")
    df = load_sheet(SHEETS["Budget"])

    if not df.empty and len(df.columns)>=3:
        try:
            df_n = df.copy()
            for c in [df.columns[2], df.columns[3]]:
                df_n[c] = pd.to_numeric(df_n[c], errors='coerce').fillna(0)
            tot_e = df_n[df.columns[2]].sum(); tot_d = df_n[df.columns[3]].sum()
            reste = tot_e - tot_d; pct_d = int(tot_d/tot_e*100) if tot_e>0 else 0
            m1,m2,m3,m4 = st.columns(4)
            m1.metric("💰 Budget estimé", f"{int(tot_e):,} FCFA")
            m2.metric("✅ Acomptes versés", f"{int(tot_d):,} FCFA")
            m3.metric("⏳ Reste à payer", f"{int(reste):,} FCFA")
            m4.metric("📊 Consommé", f"{pct_d}%")
        except Exception: pass

    tab1b, tab2b = st.tabs(["📊 Tableau", "✏️ Modifier"])
    with tab1b:
        if not df.empty: st.dataframe(df.fillna(""), use_container_width=True, hide_index=True)
    with tab2b:
        if not df.empty:
            postes = [safe(r.iloc[1]) for _, r in df.iterrows() if safe(r.iloc[1])]
            sel_p = st.selectbox("Poste", postes)
            idx_p = next((i for i, r in df.iterrows() if safe(r.iloc[1]) == sel_p), None)
            if idx_p is not None:
                rp = df.iloc[idx_p]
                with st.form("edit_budget"):
                    b1,b2,b3 = st.columns(3)
                    est = b1.number_input("Budget estimé (FCFA)", min_value=0.0,
                        value=float(str(rp.iloc[2]).replace(",","")) if safe(rp.iloc[2]) else 0.0)
                    dep = b2.number_input("Acompte versé (FCFA)", min_value=0.0,
                        value=float(str(rp.iloc[3]).replace(",","")) if safe(rp.iloc[3]) else 0.0)
                    paye_opts=["Non","Partiellement","Oui"]
                    cur_p = safe(rp.iloc[5]); p_i = paye_opts.index(cur_p) if cur_p in paye_opts else 0
                    paye = b3.selectbox("Payé ?", paye_opts, index=p_i)
                    b4,b5 = st.columns(2)
                    fourn = b4.text_input("Fournisseur", value=safe(rp.iloc[6]) if len(rp)>6 else "")
                    notes = b5.text_input("Notes", value=safe(rp.iloc[7]) if len(rp)>7 else "")
                    if st.form_submit_button("💾 Enregistrer"):
                        df = force_str_cols(df, list(df.columns))
                        df.iat[idx_p,2]=est; df.iat[idx_p,3]=dep; df.iat[idx_p,4]=est-dep; df.iat[idx_p,5]=paye
                        if len(df.columns)>6: df.iat[idx_p,6]=fourn
                        if len(df.columns)>7: df.iat[idx_p,7]=notes
                        save_sheet(SHEETS["Budget"],df); st.success("✅ Mis à jour !"); st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: INVITÉS
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🎉 Invités":
    hero()
    section("🎉 Gestion des Invités")
    df = load_sheet(SHEETS["Invites"])

    if not df.empty:
        df_c = df.dropna(how='all')
        total_i = len(df_c)
        rsvp_col = next((c for c in df_c.columns if "rsvp" in str(c).lower()), None)
        conf_i = df_c[rsvp_col].astype(str).str.contains("oui|confirmé|présent", case=False, na=False).sum() if rsvp_col else 0
        i1,i2,i3 = st.columns(3)
        i1.metric("Total invités", total_i); i2.metric("RSVP confirmés", conf_i); i3.metric("En attente", total_i-conf_i)

    tab1i, tab2i = st.tabs(["📋 Liste", "➕ Ajouter"])
    with tab1i:
        if not df.empty: st.dataframe(df.fillna(""), use_container_width=True, hide_index=True)
        else: st.info("Aucun invité enregistré.")
    with tab2i:
        with st.form("add_invite"):
            i_a1,i_a2 = st.columns(2)
            nom_i=i_a1.text_input("Nom & Prénom *"); cote_i=i_a2.selectbox("Côté",["Côté Marié","Côté Mariée","Les deux"])
            i_b1,i_b2 = st.columns(2)
            tel_i=i_b1.text_input("Téléphone"); email_i=i_b2.text_input("Email")
            i_c1,i_c2,i_c3 = st.columns(3)
            tab_n=i_c1.text_input("Table N°"); menu_i=i_c2.selectbox("Menu",["Standard","Végétarien","Sans gluten","Enfant"])
            rsvp_i=i_c3.selectbox("RSVP",["En attente","Confirmé ✅","Décliné ❌"])
            i_d1,i_d2,i_d3 = st.columns(3)
            allergie=i_d1.text_input("Allergies"); transport=i_d2.text_input("Transport"); hebergement=i_d3.selectbox("Hébergement",["Non","Oui - à organiser","Oui - organisé"])
            notes_i=st.text_input("Notes")
            if st.form_submit_button("➕ Ajouter"):
                if nom_i:
                    n_num = len(df.dropna(how='all'))+1 if not df.empty else 1
                    nr=pd.DataFrame([[n_num,nom_i,cote_i,tel_i,email_i,tab_n,menu_i,allergie,rsvp_i,transport,hebergement,notes_i]],
                                    columns=df.columns[:12] if len(df.columns)>=12 else list(range(12)))
                    df=pd.concat([df,nr],ignore_index=True)
                    save_sheet(SHEETS["Invites"],df); st.success(f"✅ {nom_i} ajouté(e) !"); st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: FOURNISSEURS
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🤝 Fournisseurs":
    hero()
    section("🤝 Fournisseurs & Prestataires")
    df = load_sheet(SHEETS["Fournisseurs"])

    tab1f, tab2f = st.tabs(["📋 Liste fournisseurs", "➕ Ajouter / Modifier"])

    with tab1f:
        if not df.empty:
            for _, row in df.iterrows():
                cat    = safe(row.iloc[0]); nom_f  = safe(row.iloc[1]) or "Non défini"
                svc    = safe(row.iloc[2]); tel_f  = safe(row.iloc[3])
                email_f= safe(row.iloc[4]); montant= safe(row.iloc[6])
                contrat= safe(row.iloc[9]) if len(row)>9 else "Non"
                contrat_ok = "oui" in contrat.lower()
                bdg_f = f'<span class="badge-done">✅ Contrat signé</span>' if contrat_ok else f'<span class="badge-pending">⏳ Contrat non signé</span>'
                contact_line = "&nbsp;|&nbsp;".join(filter(None,[
                    f"📞 {tel_f}" if tel_f else "",
                    f"✉️ {email_f}" if email_f else "",
                    f"💰 {montant} FCFA" if montant and montant not in ("0","") else "",
                ]))
                st.markdown(f"""
                <div style='background:white;border-radius:10px;padding:14px 18px;margin-bottom:10px;
                  box-shadow:0 2px 8px rgba(0,0,0,0.06);border-left:4px solid {CHAMPAGNE};'>
                  <div style='display:flex;justify-content:space-between;align-items:start;'>
                    <div>
                      <div style='font-weight:700;color:{SAUGE_DARK};font-size:1rem;'>{cat} — {nom_f}</div>
                      <div style='color:#555;font-size:0.88rem;margin-top:3px;'>{svc}</div>
                      <div style='font-size:0.82rem;color:#777;margin-top:5px;'>{contact_line}</div>
                    </div>
                    <div style='margin-left:12px;'>{bdg_f}</div>
                  </div>
                </div>""", unsafe_allow_html=True)
        else:
            st.info("Aucun fournisseur enregistré.")

    with tab2f:
        st.markdown("### ➕ Ajouter un fournisseur")
        with st.form("add_fourn"):
            f1,f2 = st.columns(2)
            cat_f=f1.selectbox("Catégorie",["Traiteur","Photo/Vidéo","Musique","Décoration","Beauté","Gâteau","Transport","Papeterie","Hébergement","Autre"])
            nom_ff=f2.text_input("Nom société / Personne *")
            svc_f=st.text_input("Service fourni")
            f3,f4 = st.columns(2)
            tel_ff=f3.text_input("Téléphone"); email_ff=f4.text_input("Email")
            adresse_f=st.text_input("Adresse")
            f5,f6,f7 = st.columns(3)
            mont_f=f5.number_input("Montant contrat (FCFA)",min_value=0.0)
            acomp_f=f6.number_input("Acompte versé (FCFA)",min_value=0.0)
            cont_f=f7.selectbox("Contrat signé ?",["Non","Oui"])
            notes_f=st.text_input("Notes")
            if st.form_submit_button("➕ Ajouter"):
                if nom_ff:
                    reste_f=mont_f-acomp_f
                    nr=pd.DataFrame([[cat_f,nom_ff,svc_f,tel_ff,email_ff,adresse_f,mont_f,acomp_f,reste_f,cont_f,notes_f]],
                                    columns=df.columns[:11] if len(df.columns)>=11 else list(range(11)))
                    df=pd.concat([df,nr],ignore_index=True)
                    save_sheet(SHEETS["Fournisseurs"],df); st.success(f"✅ {nom_ff} ajouté(e) !"); st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: NOTES PCO
# ═══════════════════════════════════════════════════════════════════════════
elif page == "📝 Notes PCO":
    hero()
    section("📝 Journal PCO — Notes & Décisions")
    df = load_sheet(SHEETS["Notes"])

    tab1n, tab2n = st.tabs(["📖 Journal", "➕ Ajouter une note"])

    with tab1n:
        if not df.empty:
            for _, row in df.iterrows():
                date_n=safe(row.iloc[0]); sujet=safe(row.iloc[1]); note=safe(row.iloc[2])
                action=safe(row.iloc[3]); resp=safe(row.iloc[4]); statut_n=safe(row.iloc[5]); prio=safe(row.iloc[6])
                border_n = SAUGE_DARK if "urgent" in prio.lower() else CHAMPAGNE
                _, bdg_n = badge(statut_n)
                action_html = f'<div style="font-size:0.83rem;color:{CHAMP_DARK};margin-top:5px;">➡️ {action} ({resp})</div>' if action else ""
                st.markdown(f"""
                <div style='background:white;border-left:5px solid {border_n};border-radius:10px;
                  padding:13px 18px;margin-bottom:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);'>
                  <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;'>
                    <div>
                      <span style='color:#999;font-size:0.8rem;'>📅 {date_n}</span>
                      <span style='margin-left:10px;font-weight:700;color:{SAUGE_DARK};'>{sujet}</span>
                      <span style='margin-left:10px;font-size:0.78rem;color:#888;'>{prio}</span>
                    </div>
                    <div>{bdg_n}</div>
                  </div>
                  <div style='color:#444;font-size:0.88rem;'>{note}</div>
                  {action_html}
                </div>""", unsafe_allow_html=True)
        else:
            st.info("Aucune note PCO pour l'instant.")

    with tab2n:
        with st.form("add_note"):
            n1,n2 = st.columns(2)
            date_note=n1.date_input("Date",value=date.today()); sujet_note=n2.text_input("Sujet *")
            note_text=st.text_area("Note / Décision *",height=90)
            n3,n4 = st.columns(2)
            action_note=n3.text_input("Action requise"); resp_note=n4.text_input("Responsable")
            n5,n6 = st.columns(2)
            st_note=n5.selectbox("Statut",["À faire","En cours","Fait"])
            prio_note=n6.selectbox("Priorité",["🔴 Urgent","🟡 Normal","🟢 Faible"])
            if st.form_submit_button("💾 Enregistrer"):
                if sujet_note and note_text:
                    nr=pd.DataFrame([[date_note.strftime("%d/%m/%Y"),sujet_note,note_text,action_note,resp_note,st_note,prio_note]],
                                    columns=df.columns[:7] if len(df.columns)>=7 else list(range(7)))
                    df=pd.concat([df,nr],ignore_index=True)
                    save_sheet(SHEETS["Notes"],df); st.success("✅ Note enregistrée !"); st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Configuration":
    hero()
    section("⚙️ Configuration du Mariage")
    config = get_config()

    col1c, col2c = st.columns(2)
    with col1c:
        st.markdown("### 👰🤵 Les mariés & infos clés")
        with st.form("config_form"):
            marie    = st.text_input("Prénom du marié",  value=config.get("Marié","Hussen Junior"))
            mariee   = st.text_input("Prénom de la mariée", value=config.get("Mariée","Fransesca"))
            date_m   = st.text_input("Date du mariage",  value=config.get("Date du mariage","01/08/2025"))
            pco_name = st.text_input("Nom du PCO",       value=config.get("PCO",""))
            nb_inv   = st.text_input("Nombre total invités", value=config.get("Nombre invités total",""))
            budget_t = st.text_input("Budget total (FCFA)",  value=config.get("Budget total estimé",""))
            lieu_c   = st.text_input("Lieu cérémonie",   value=config.get("Lieu cérémonie",""))
            lieu_r   = st.text_input("Lieu réception",   value=config.get("Lieu réception",""))
            if st.form_submit_button("💾 Enregistrer"):
                save_config({"Marié":marie,"Mariée":mariee,"Date du mariage":date_m,"PCO":pco_name,
                    "Nombre invités total":nb_inv,"Budget total estimé":budget_t,
                    "Lieu cérémonie":lieu_c,"Lieu réception":lieu_r})
                st.success("✅ Configuration enregistrée !")

    with col2c:
        st.markdown("### 🎨 Palette officielle — Sauge & Champagne")
        swatches_c = [
            (SAUGE_DARK,"Vert Sauge Foncé",SAUGE_DARK),
            (SAUGE,"Vert Sauge Clair",SAUGE_DARK),
            (CHAMPAGNE,"Champagne Satiné",CHAMP_DARK),
            (IVORY,"Ivoire","#888"),
        ]
        ccc = st.columns(4)
        for i,(bg,name,txt) in enumerate(swatches_c):
            ccc[i].markdown(f"""<div style='text-align:center;'>
              <div style='background:{bg};width:64px;height:64px;border-radius:50%;margin:auto;border:2px solid #ccc;'></div>
              <div style='font-size:0.75rem;margin-top:5px;color:{txt};font-weight:600;'>{name}</div>
              <div style='font-size:0.7rem;color:#aaa;'>{bg}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📥 Fichier Excel")
        st.info(f"📁 **{EXCEL_FILE}**")
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE,"rb") as f:
                st.download_button("⬇️ Télécharger Excel",data=f,file_name=EXCEL_FILE,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### 🔄 Réinitialisation")
        if st.button("🗑️ Réinitialiser toutes les données"):
            if st.session_state.get("confirm_reset"):
                if os.path.exists(EXCEL_FILE): os.remove(EXCEL_FILE)
                init_excel()
                st.success("✅ Fichier réinitialisé !"); st.session_state["confirm_reset"]=False; st.rerun()
            else:
                st.session_state["confirm_reset"]=True
                st.warning("⚠️ Cliquez à nouveau pour confirmer la réinitialisation.")