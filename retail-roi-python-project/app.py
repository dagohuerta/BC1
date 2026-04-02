from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

from retail_roi_model.engine import WorkbookDrivenRetailROIModel, outputs_to_jsonable, npv, irr, payback_period

import io
import matplotlib.pyplot as plt
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
except ImportError:
    pass

st.set_page_config(page_title="Strategic ROI Analysis", layout="wide")

def format_usd(val):
    if val is None: return "N/D"
    return f"${val:,.0f}" if abs(val) >= 1 else f"${val:,.2f}"

def format_m(val):
    if val is None: return "N/D"
    if abs(val) >= 1_000_000:
        return f"${val/1_000_000:,.1f} M"
    return format_usd(val)

def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Cargar Estilos Externos
try:
    local_css("style.css")
except FileNotFoundError:
    pass

# Mapeo de Racionales Profesionales
RATIONALE_MAP = {
    "sales": "Optimización de ingresos mediante mejor disponibilidad y previsión operativa coordinada.",
    "inventory_reduction": "Reducción de capital inmovilizado y mejora en eficiencia de stock mediante visibilidad.",
    "margin": "Protección y expansión de margen mediante optimización estratégica de precios y promociones.",
    "labor": "Aumento de productividad del personal y eficiencia en operaciones de tienda/almacén.",
    "logistics": "Optimización de costos logísticos y de transporte mediante una red inteligente."
}

def get_total_manual_investment():
    # Sumar inversiones hasta el horizonte seleccionado
    horizon = st.session_state.get('manual_horizon', 5)
    total = 0.0
    for i in range(1, horizon + 1):
        inv = st.session_state.annual_investments.get(i, {'software': 0, 'impl': 0, 'extra': 0})
        total += inv['software'] + inv['impl'] + inv['extra']
    return total

# --- Función para mostrar el Reporte Ejecutivo ---
@st.dialog("Configuración de Inversiones Anuales", width="large")
def configure_investments():
    horizon = st.session_state.get('manual_horizon', 5)
    st.write(f"Ingresa los montos de inversión proyectados para el horizonte de **{horizon} años**.")
    
    # Encabezados
    c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
    c1.write("**Año**")
    c2.write("**Licencias**")
    c3.write("**Implementación**")
    c4.write("**Adicional / Otros**")
    
    for i in range(1, horizon + 1):
        cols = st.columns([1, 2, 2, 2])
        cols[0].write(f"Año {i}")
        st.session_state.annual_investments[i]['software'] = cols[1].number_input(f"Software Y{i}", min_value=0.0, value=st.session_state.annual_investments[i]['software'], label_visibility="collapsed", key=f"soft_{i}")
        st.session_state.annual_investments[i]['impl'] = cols[2].number_input(f"Impl Y{i}", min_value=0.0, value=st.session_state.annual_investments[i]['impl'], label_visibility="collapsed", key=f"impl_{i}")
        st.session_state.annual_investments[i]['extra'] = cols[3].number_input(f"Extra Y{i}", min_value=0.0, value=st.session_state.annual_investments[i]['extra'], label_visibility="collapsed", key=f"extra_{i}")
    
    if st.button("Guardar y Cerrar", type="primary", use_container_width=True):
        st.rerun()

@st.dialog("Reporte Ejecutivo de ROI", width="large")
def show_executive_report(results):
    # Definir paleta Dark Mode Neon
    neon_palette = ['#00a3ff', '#00e676', '#ffc107', '#ff5252', '#bb86fc', '#03dac6', '#cf6679']

    st.markdown(f'<div class="header-monday">Reporte Ejecutivo de ROI</div>', unsafe_allow_html=True)
    st.markdown(f"**Análisis Estratégico Obsidian:** {results['cliente']} | {results['retailer_type']}")
    
    st.write("---")
    
    # Métricas Principales en Cards Dark Elite
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f"""<div class="monday-card"><div class="monday-card-label">VPN (NPV)</div><div class="monday-card-value">{format_m(results['npv'])}</div></div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""<div class="monday-card card-green"><div class="monday-card-label">TIR (IRR)</div><div class="monday-card-value">{results['irr'] if results['irr'] is not None else 'N/D'}%</div></div>""", unsafe_allow_html=True)
    with m3:
        st.markdown(f"""<div class="monday-card card-amber"><div class="monday-card-label">Payback</div><div class="monday-card-value">{results['payback'] or 'N/D'}</div></div>""", unsafe_allow_html=True)
    with m4:
        total_b = sum(results['total_benefit'])
        st.markdown(f"""<div class="monday-card card-purple"><div class="monday-card-label">Total Beneficios</div><div class="monday-card-value">{format_m(total_b)}</div></div>""", unsafe_allow_html=True)

    st.write("---")

    # Layout de Gráficos (Dark Theme)
    g1, g2 = st.columns([2, 1])
    
    with g1:
        st.write("**Trayectoria de Beneficios**")
        years_list = list(range(1, results['years'] + 1))
        fig_traj = go.Figure()
        fig_traj.add_trace(go.Bar(x=years_list, y=results['total_benefit'], name='Beneficios', marker_color=neon_palette[1])) # Neon Green
        fig_traj.add_trace(go.Bar(x=years_list, y=results['total_investment'], name='Inversiones', marker_color='#444444')) # Stealth Gray
        fig_traj.add_trace(go.Scatter(x=years_list, y=results['total_cumulative'], name='Flujo Acumulado', line=dict(color=neon_palette[0], width=5))) # Neon Blue
        
        fig_traj.update_layout(template='plotly_dark', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=400, margin=dict(l=0, r=0, t=20, b=0), legend=dict(orientation="h", y=1.2))
        fig_traj.update_yaxes(tickprefix="$", tickformat=",.0f", gridcolor="#333")
        st.plotly_chart(fig_traj, width="stretch")

    with g2:
        if results['selection']:
            st.write("**Desglose de Valor**")
            benefits_by_mod = [sum(results['module_results'][m]['benefit']) for m in results['selection']]
            
            # Pie Chart de Alto Contraste Neon
            fig_pie = px.pie(values=benefits_by_mod, names=results['selection'], hole=0.5)
            # Colores Neón variados
            pie_colors = [neon_palette[0], neon_palette[4], neon_palette[5], neon_palette[2], neon_palette[1], neon_palette[3]]
            fig_pie.update_traces(marker=dict(colors=pie_colors), textinfo='percent', hovertemplate='%{label}<br>$%{value:,.0f}', textfont_size=14)
            fig_pie.update_layout(template='plotly_dark', paper_bgcolor='rgba(0,0,0,0)', showlegend=True, legend=dict(orientation="h", y=-0.2), margin=dict(l=0, r=0, t=0, b=0), height=400)
            st.plotly_chart(fig_pie, width="stretch")

    with st.expander("📝 Detalle Técnico de Beneficios y Racionales"):
        st.write("Análisis detallado de la lógica de negocio aplicada por módulo.")
        benefit_rows = []
        for mod_name in results['selection']:
            mod_data = results['module_results'].get(mod_name, {})
            # Derivar racionales y porcentajes
            impacts = []
            if 'aspect_pcts' in results and mod_name in results['aspect_pcts']:
                # Caso Manual: tenemos los porcentajes exactos
                for aspect, pct in results['aspect_pcts'][mod_name].items():
                    impacts.append(f"{aspect.replace('_',' ').title()}: {pct}%")
            else:
                # Caso Excel o genérico
                impacts.append("Impacto según modelo estratégico")
            
            # Racional combinado basado en perfiles
            racional = "Calculado según parámetros de la industria para retail."
            if mod_name in st.session_state.module_profiles:
                racional = " ".join([RATIONALE_MAP.get(a, "") for a in st.session_state.module_profiles[mod_name]])
            
            # El monto es la suma de los beneficios anuales para este módulo
            total_monto = sum(mod_data.get('benefit', [0]))
            
            benefit_rows.append({
                "Módulo / Beneficio": mod_name,
                "Justificación (Racional)": racional,
                "Porcentaje Aplicado": ", ".join(impacts),
                "Monto Calculado (Total)": format_usd(total_monto)
            })
        
        df_benefits = pd.DataFrame(benefit_rows)
        st.table(df_benefits)

    with st.expander("📊 Ver Detalle Anual Proyectado"):
        st.write("Cifras proyectadas año con año (métricas transpuestas)")
        # Lógica de construcción de tabla
        years_list = list(range(1, results['years'] + 1))
        detail_df = pd.DataFrame({
            'Beneficio': results['total_benefit'],
            'Inversión': results['total_investment'],
            'Flujo neto': results['total_cashflow'],
            'Cumulado': results['total_cumulative']
        }, index=[f"Año {i}" for i in years_list])
        
        df_t = detail_df.T
        df_t['Total'] = df_t.sum(axis=1)
        st.table(df_t.map(format_usd))

    # --- Exportación PPTX Ejecutiva Premium ---
    try:
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.enum.text import PP_ALIGN
        
        prs = Presentation()
        # Definir Colores Corporativos
        COLOR_BLUE = RGBColor(0, 163, 255)  # Neon Blue
        COLOR_GREEN = RGBColor(0, 230, 118) # Neon Green
        COLOR_DARK = RGBColor(14, 17, 23)   # Dark Background
        COLOR_WHITE = RGBColor(255, 255, 255)
        
        # 1. Slide de Portada (Executive Cover)
        slide_cover = prs.slides.add_slide(prs.slide_layouts[6])
        # Fondo degradado simulado
        rect = slide_cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        rect.fill.solid()
        rect.fill.fore_color.rgb = COLOR_DARK
        rect.line.fill.background()
        
        # Título Principal
        title_box = slide_cover.shapes.add_textbox(Inches(0.5), Inches(3), Inches(9), Inches(1.5))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = "PROYECCIÓN ESTRATÉGICA DE ROI"
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER
        
        # Subtítulo (Cliente)
        sub_box = slide_cover.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(1))
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"Análisis Preparado para: {results['cliente']} | {results['retailer_type']}"
        p.font.size = Pt(20)
        p.font.color.rgb = COLOR_BLUE
        p.alignment = PP_ALIGN.CENTER

        # 2. Resumen Financiero Ejecutivo (KPI Dashboard)
        slide_kpi = prs.slides.add_slide(prs.slide_layouts[6])
        # Header
        header = slide_kpi.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1))
        header.fill.solid()
        header.fill.fore_color.rgb = COLOR_BLUE
        header.line.fill.background()
        
        title = slide_kpi.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(8), Inches(0.6))
        tf = title.text_frame
        p = tf.paragraphs[0]
        p.text = "Métricas Clave de Retorno (Executive Summary)"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        
        # KPI Boxes (4 column layout)
        kpi_list = [
            ("NPV (Valor Presente)", format_m(results['npv'])),
            ("IRR (Tasa Interna)", f"{results['irr']}%"),
            ("Payback (Recuperación)", results['payback']),
            ("Beneficio Total", format_m(sum(results['total_benefit'])))
        ]
        
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(2.1)
        height = Inches(1.2)
        
        for name, val in kpi_list:
            # Shape for background
            box = slide_kpi.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(30, 30, 30)
            box.line.color.rgb = COLOR_BLUE
            
            # Label
            lbl = slide_kpi.shapes.add_textbox(left, top + Inches(0.15), width, Inches(0.3))
            tf = lbl.text_frame
            p = tf.paragraphs[0]
            p.text = name
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_BLUE
            p.alignment = PP_ALIGN.CENTER
            
            # Value
            vbox = slide_kpi.shapes.add_textbox(left, top + Inches(0.5), width, Inches(0.5))
            tf = vbox.text_frame
            p = tf.paragraphs[0]
            p.text = val
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = COLOR_WHITE
            p.alignment = PP_ALIGN.CENTER
            left += Inches(2.25)

        # 3. Trayectoria de Beneficios (Financial Trajectory)
        slide_traj = prs.slides.add_slide(prs.slide_layouts[6])
        # Background
        bg = slide_traj.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_DARK
        
        # Chart logic
        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(10, 5))
        years = list(range(1, results['years'] + 1))
        # Stylized Bars & Lines
        ax.bar(years, results['total_benefit'], color='#00e676', alpha=0.8, label='Beneficios Anuales')
        ax.bar(years, results['total_investment'], color='#ff5252', alpha=0.8, label='Inversiones')
        ax2 = ax.twinx()
        ax2.plot(years, results['total_cumulative'], color='#00a3ff', linewidth=4, marker='o', markersize=10, label='Flujo Acumulado')
        
        ax.set_title("Evolución de Flujo de Caja y ROI", color='white', fontsize=16, pad=20)
        ax.legend(loc='upper left', facecolor='#1e1e1e', edgecolor='#333')
        ax.grid(axis='y', color='#333', linestyle='--')
        
        img_buf = io.BytesIO()
        fig.savefig(img_buf, format='png', bbox_inches='tight', dpi=150)
        img_buf.seek(0)
        plt.close(fig)
        
        slide_traj.shapes.add_picture(img_buf, Inches(0.5), Inches(1.5), width=Inches(9))

        # 4. Detalle de Racionales (Business Case Details)
        slide_rac = prs.slides.add_slide(prs.slide_layouts[6])
        # White background for readable data
        title_box = slide_rac.shapes.add_textbox(Inches(0.2), Inches(0.2), Inches(9), Inches(0.6))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = "Justificación Estratégica por Módulo"
        p.font.size = Pt(28)
        p.font.bold = True
        
        # Table
        rows = len(results['selection']) + 1
        cols = 3
        table = slide_rac.shapes.add_table(rows, cols, Inches(0.5), Inches(1.2), Inches(9), Inches(4)).table
        
        # Headers
        table.cell(0, 0).text = "Módulo Oracle"
        table.cell(0, 1).text = "Principales Palancas de Valor"
        table.cell(0, 2).text = "Beneficio Total"
        
        for k in range(3):
            table.cell(0, k).fill.solid()
            table.cell(0, k).fill.fore_color.rgb = COLOR_BLUE
            table.cell(0, k).text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE
            table.cell(0, k).text_frame.paragraphs[0].font.bold = True

        for idx, mod_name in enumerate(results['selection']):
            mod_data = results['module_results'].get(mod_name, {})
            # Derivar aspectos
            impacts = []
            if 'aspect_pcts' in results and mod_name in results['aspect_pcts']:
                impacts = [f"{a.replace('_',' ').title()}" for a, p in results['aspect_pcts'][mod_name].items()]
            
            table.cell(idx+1, 0).text = mod_name
            table.cell(idx+1, 1).text = ", ".join(impacts)
            table.cell(idx+1, 2).text = format_usd(sum(mod_data.get('benefit', [0])))

        # 5. Proyecciones Anualizadas (Full Data)
        slide_table = prs.slides.add_slide(prs.slide_layouts[6])
        title_box = slide_table.shapes.add_textbox(Inches(0.2), Inches(0.2), Inches(9), Inches(0.6))
        title_box.text_frame.paragraphs[0].text = "Proyecciones Financieras Anualizadas ($ USD)"
        title_box.text_frame.paragraphs[0].font.size = Pt(24)
        
        # Transposed table
        detail_rows = ["Año "+str(i) for i in range(1, results['years'] + 1)]
        detail_rows.append("Total")
        
        cols = results['years'] + 2 # Metrics + Years + Total
        rows = 5 # Benefit, Inv, Net, Cum
        table = slide_table.shapes.add_table(5, results['years'] + 2, Inches(0.2), Inches(1.5), Inches(9.6), Inches(2.5)).table
        
        metrics = ["Año", "Beneficio", "Inversión", "Flujo Neto", "Cumulado"]
        for r, m in enumerate(metrics):
            table.cell(r, 0).text = m
            table.cell(r, 0).fill.solid()
            table.cell(r, 0).fill.fore_color.rgb = RGBColor(240, 240, 240)
            table.cell(r, 0).text_frame.paragraphs[0].font.bold = True
            
        for y in range(1, results['years']+1):
            table.cell(0, y).text = f"Y{y}"
            table.cell(1, y).text = format_usd(results['total_benefit'][y-1])
            table.cell(2, y).text = format_usd(results['total_investment'][y-1])
            table.cell(3, y).text = format_usd(results['total_cashflow'][y-1])
            table.cell(4, y).text = format_usd(results['total_cumulative'][y-1])
            
        # Total Column
        table.cell(0, results['years']+1).text = "Total"
        table.cell(1, results['years']+1).text = format_usd(sum(results['total_benefit']))
        table.cell(2, results['years']+1).text = format_usd(sum(results['total_investment']))
        table.cell(3, results['years']+1).text = format_usd(sum(results['total_cashflow']))
        table.cell(4, results['years']+1).text = "N/A"

        # Guardar en memoria
        pptx_buf = io.BytesIO()
        prs.save(pptx_buf)
        pptx_buf.seek(0)

        st.write("---")
        c1, c2 = st.columns([3, 1])
        with c2:
            st.download_button(
                "📈 DESCARGAR PPTX EJECUTIVO",
                data=pptx_buf,
                file_name=f"ROI_Plan_Estrategico_{results['cliente']}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                help="Generar el Business Case completo en formato PowerPoint profesional"
            )
    except Exception as e:
        import traceback
        st.warning(f"Exportación PPTX no disponible: {e}")
        st.caption(traceback.format_exc())

    st.info("Este reporte resume los beneficios proyectados basados en los parámetros ingresados y las curvas de adopción estándar de la industria.")

# Rangos por aspecto - inicializar en session_state
if 'aspect_ranges' not in st.session_state:
    st.session_state.aspect_ranges = {
        "sales": (0, 20),
        "inventory_reduction": (0, 15),
        "margin": (0, 10),
        "labor": (0, 12),
        "logistics": (0, 8),
    }

# Inicializar inversiones anuales en session_state si no existe
if 'annual_investments' not in st.session_state:
    st.session_state.annual_investments = {
        i: {'software': 1200000.0 if i==1 else 0.0, 'impl': 800000.0 if i==1 else 0.0, 'extra': 0.0} 
        for i in range(1, 6)
    }

aspect_ranges = st.session_state.aspect_ranges

# Inicializar session_state para configuraciones
if 'module_options' not in st.session_state:
    st.session_state.module_options = [
        "Inventory Optimization", "Pricing", "Merchandising", "Customer Experience",
        "Supply Chain", "Retail Insights", "Store Operations", "Loyalty", "Omnichannel", "Data Science"
    ]

if 'module_profiles' not in st.session_state:
    st.session_state.module_profiles = {
        "Inventory Optimization": ["inventory_reduction"],
        "Pricing": ["margin"],
        "Merchandising": ["sales"],
        "Customer Experience": ["sales", "margin"],
        "Supply Chain": ["logistics"],
        "Retail Insights": ["sales", "inventory_reduction"],
        "Store Operations": ["labor"],
        "Loyalty": ["sales"],
        "Omnichannel": ["sales", "logistics"],
        "Data Science": ["margin", "inventory_reduction"],
    }

if 'benefit_params' not in st.session_state:
    st.session_state.benefit_params = {}
    for module in st.session_state.module_options:
        st.session_state.benefit_params[module] = {}
        for aspect in st.session_state.module_profiles[module]:
            min_val, max_val = aspect_ranges[aspect]
            st.session_state.benefit_params[module][aspect] = {"min": min_val + 1, "max": max_val - 1}

st.title("Retail ROI Model")

main_tab, config_tab = st.tabs(["Cálculo", "Configuración"])

with config_tab:
    config_subtab1, config_subtab2, config_subtab3 = st.tabs(["Módulos Oracle", "Parámetros de Beneficio", "Aspectos"])

    with config_subtab1:
        st.subheader("Gestión de Módulos Oracle")
        
        # --- Toolbar Dual-View (Agregar + Buscar + Vista) ---
        with st.container():
            t1, t2, t3 = st.columns([2, 2, 1])
            with t1:
                new_module = st.text_input("Nombre del módulo", key="new_module_compact", label_visibility="collapsed", placeholder="➕ Nuevo módulo Oracle...")
                if new_module and new_module not in st.session_state.module_options:
                    if st.button("Crear Módulo", key="btn_create_dual"):
                        st.session_state.module_options.append(new_module)
                        st.session_state.module_profiles[new_module] = []
                        st.session_state.benefit_params[new_module] = {}
                        st.rerun()
            with t2:
                search_query = st.text_input("🔍 Buscar...", key="module_search_dual", label_visibility="collapsed", placeholder="🔍 Buscar módulos...")
            with t3:
                view_mode = st.radio("Vista", ["⊞", "☰"], index=1, horizontal=True, label_visibility="collapsed")

        # Mapeo de modo de vista
        grid_mode = (view_mode == "⊞")
        filtered_modules = [m for m in st.session_state.module_options if search_query.lower() in m.lower()]
        
        if not filtered_modules:
            st.warning("No se encontraron resultados.")
        else:
            if grid_mode:
                # Grid Optimizado (3 columnas)
                cols_per_row = 3
                for i in range(0, len(filtered_modules), cols_per_row):
                    row_modules = filtered_modules[i:i + cols_per_row]
                    cols = st.columns(cols_per_row)
                    for j, module in enumerate(row_modules):
                        with cols[j]:
                            with st.container(border=True):
                                n_aspects = len(st.session_state.module_profiles.get(module, []))
                                st.markdown(f"**{module}** <small style='color:var(--dark-text-dim); float:right;'>{n_aspects} asp.</small>", unsafe_allow_html=True)
                                
                                current_aspects = st.session_state.module_profiles.get(module, [])
                                selected_aspects = st.multiselect("Impacto", list(aspect_ranges.keys()), default=current_aspects, key=f"aspects_grid_{module}", label_visibility="collapsed")
                                
                                if selected_aspects != current_aspects:
                                    st.session_state.module_profiles[module] = selected_aspects
                                    for aspect in selected_aspects:
                                        if aspect not in st.session_state.benefit_params[module]:
                                            min_v, max_v = aspect_ranges[aspect]
                                            st.session_state.benefit_params[module][aspect] = {"min": min_v + 1, "max": max_v - 1}
                                    for a in list(st.session_state.benefit_params[module].keys()):
                                        if a not in selected_aspects:
                                            del st.session_state.benefit_params[module][a]
                                    st.rerun()

                                with st.expander("Opciones", expanded=False):
                                    new_name = st.text_input("Nombre", module, key=f"rename_grid_{module}", label_visibility="collapsed")
                                    if new_name != module and new_name.strip() != "":
                                        idx = st.session_state.module_options.index(module)
                                        st.session_state.module_options[idx] = new_name
                                        st.session_state.module_profiles[new_name] = st.session_state.module_profiles.pop(module)
                                        st.session_state.benefit_params[new_name] = st.session_state.benefit_params.pop(module)
                                        st.rerun()
                                    if st.button("🗑️ Eliminar", key=f"del_grid_{module}", use_container_width=True):
                                        st.session_state.module_options.remove(module)
                                        del st.session_state.module_profiles[module]
                                        del st.session_state.benefit_params[module]
                                        st.rerun()
            else:
                # Vista de Lista (Filas compactas)
                for module in filtered_modules:
                    with st.container(border=True):
                        c1, c2, c3 = st.columns([2, 3, 1])
                        with c1:
                            st.markdown(f"**{module}**")
                            n_aspects = len(st.session_state.module_profiles.get(module, []))
                            st.caption(f"{n_aspects} aspectos seleccionados")
                        with c2:
                            current_aspects = st.session_state.module_profiles.get(module, [])
                            selected_aspects = st.multiselect("Impacto", list(aspect_ranges.keys()), default=current_aspects, key=f"aspects_list_{module}", label_visibility="collapsed")
                            if selected_aspects != current_aspects:
                                st.session_state.module_profiles[module] = selected_aspects
                                for aspect in selected_aspects:
                                    if aspect not in st.session_state.benefit_params[module]:
                                        min_v, max_v = aspect_ranges[aspect]
                                        st.session_state.benefit_params[module][aspect] = {"min": min_v + 1, "max": max_v - 1}
                                for a in list(st.session_state.benefit_params[module].keys()):
                                    if a not in selected_aspects:
                                        del st.session_state.benefit_params[module][a]
                                st.rerun()
                        with c3:
                            with st.popover("⚙️"):
                                new_name = st.text_input("Renombrar módulo", module, key=f"rename_list_{module}")
                                if new_name != module and new_name.strip() != "":
                                    idx = st.session_state.module_options.index(module)
                                    st.session_state.module_options[idx] = new_name
                                    st.session_state.module_profiles[new_name] = st.session_state.module_profiles.pop(module)
                                    st.session_state.benefit_params[new_name] = st.session_state.benefit_params.pop(module)
                                    st.rerun()
                                if st.button("🗑️ Eliminar Módulo", key=f"del_list_{module}", type="secondary", use_container_width=True):
                                    st.session_state.module_options.remove(module)
                                    del st.session_state.module_profiles[module]
                                    del st.session_state.benefit_params[module]
                                    st.rerun()

    with config_subtab2:
        st.subheader("Parámetros de Beneficio por Módulo y Aspecto")
        for module in st.session_state.module_options:
            st.write(f"**{module}**")
            if module in st.session_state.module_profiles:
                for aspect in st.session_state.module_profiles[module]:
                    col_min, col_max = st.columns(2)
                    min_val, max_val = aspect_ranges[aspect]
                    with col_min:
                        current_min = st.session_state.benefit_params[module][aspect]["min"]
                        new_min = st.number_input(f"Min % {module} - {aspect}", min_val, max_val, current_min, key=f"min_{module}_{aspect}")
                        st.session_state.benefit_params[module][aspect]["min"] = new_min
                    with col_max:
                        current_max = st.session_state.benefit_params[module][aspect]["max"]
                        new_max = st.number_input(f"Max % {module} - {aspect}", min_val, max_val, current_max, key=f"max_{module}_{aspect}")
                        st.session_state.benefit_params[module][aspect]["max"] = new_max
            st.write("---")

    with config_subtab3:
        st.subheader("Configuración de Aspectos e Impacto")
        st.info("Configura los rangos de impacto (min/max) para cada aspecto de negocio.")
        
        # (El manejo de aspectos se realiza en el bloque de abajo dinámicamente)
        
        # Agregar nuevo aspecto
        st.write("### Agregar nuevo aspecto")
        col_new_name, col_new_min, col_new_max, col_add_btn = st.columns([2, 1.5, 1.5, 1])
        with col_new_name:
            new_aspect_name = st.text_input("Nombre del aspecto", key="new_aspect_name")
        with col_new_min:
            new_aspect_min = st.number_input("Rango mín (%)", min_value=0, max_value=100, value=0, key="new_aspect_min")
        with col_new_max:
            new_aspect_max = st.number_input("Rango máx (%)", min_value=0, max_value=100, value=20, key="new_aspect_max")
        with col_add_btn:
            if st.button("Agregar", key="btn_add_aspect"):
                if new_aspect_name and new_aspect_name not in st.session_state.aspect_ranges:
                    if new_aspect_max > new_aspect_min:
                        st.session_state.aspect_ranges[new_aspect_name] = (new_aspect_min, new_aspect_max)
                        st.success(f"Aspecto '{new_aspect_name}' agregado ({new_aspect_min}-{new_aspect_max}%)")
                        st.rerun()
                    else:
                        st.error("El máximo debe ser mayor que el mínimo")
                else:
                    st.error("Nombre inválido o ya existe")
        
        # Listar y editar aspectos
        st.write("### Aspectos existentes")
        for aspect_name, (min_val, max_val) in list(st.session_state.aspect_ranges.items()):
            col_name, col_min, col_max, col_actions = st.columns([2, 1.5, 1.5, 1])
            
            with col_name:
                st.text(aspect_name)
            
            with col_min:
                new_min = st.number_input(f"Min {aspect_name}", min_value=0, max_value=100, value=min_val, key=f"emin_{aspect_name}")
            
            with col_max:
                new_max = st.number_input(f"Max {aspect_name}", min_value=0, max_value=100, value=max_val, key=f"emax_{aspect_name}")
            
            with col_actions:
                if st.button("🗑️", key=f"del_aspect_{aspect_name}"):
                    # Verificar si hay módulos usando este aspecto
                    modules_using = [m for m, aspects in st.session_state.module_profiles.items() if aspect_name in aspects]
                    if modules_using:
                        st.error(f"No puedes eliminar '{aspect_name}': usado por {', '.join(modules_using)}")
                    else:
                        del st.session_state.aspect_ranges[aspect_name]
                        st.success(f"Aspecto '{aspect_name}' eliminado")
                        st.rerun()
            
            # Actualizar rango si cambió
            if (new_min, new_max) != (min_val, max_val):
                if new_max > new_min:
                    st.session_state.aspect_ranges[aspect_name] = (new_min, new_max)
                    # Sincronizar con benefit_params de todos los módulos que usan este aspecto
                    for module, aspects in st.session_state.module_profiles.items():
                        if aspect_name in aspects:
                            st.session_state.benefit_params[module][aspect_name] = {
                                "min": new_min + 1,
                                "max": new_max - 1
                            }
                else:
                    st.error(f"Max debe ser > Min para {aspect_name}")
            
            st.divider()

with main_tab:
    with st.sidebar:
        st.markdown('<div class="header-monday" style="font-size: 1.6rem; border-bottom: 5px solid var(--neon-blue);">ROI Strategist</div>', unsafe_allow_html=True)
        st.write("**Parámetros del Modelo**")
        mode = st.radio("Método de entrada", ["Carga Excel", "Entrada manual"])
        
        st.write("---")
        st.caption("v5.0.0 - Obsidian Dark UI")
        
    st.markdown('<div class="header-monday">Análisis de ROI Estratégico</div>', unsafe_allow_html=True)
    st.markdown("### <span style='color:var(--neon-amber)'>Inteligencia de Negocio</span>", unsafe_allow_html=True)
    st.info("Bienvenido. Ejecuta el análisis de retorno de inversión en el nuevo entorno Obsidian Dark.")

    if mode == "Entrada manual":
        with st.sidebar:
            st.write("---")
            st.write("**Configuración de Inversión**")
            if st.button("🛠️ Inversiones Anuales", width="stretch"):
                configure_investments()
            total_inv_setup = get_total_manual_investment()
            st.caption(f"Total Configurado: {format_usd(total_inv_setup)}")

    if mode == "Carga Excel":
        st.caption("Carga el archivo .xlsm y genera el resumen de ROI y el JSON consolidado.")
        uploaded = st.file_uploader("Workbook Excel (.xlsm)", type=["xlsm", "xlsx"])

        if uploaded is not None:
            with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded.name}") as tmp:
                tmp.write(uploaded.getbuffer())
                temp_path = Path(tmp.name)

            if st.button("Calcular ROI"):
                try:
                    model = WorkbookDrivenRetailROIModel(temp_path)
                    outputs = model.run()
                    payload = outputs_to_jsonable(outputs)
                    
                    # Preparar data del reporte
                    report_data = {
                        "cliente": payload["summary_metrics"].get("client_name", "Cliente"),
                        "retailer_type": payload["summary_metrics"].get("retailer_type", "Retailer"),
                        "npv": payload["summary_metrics"].get("npv", 0),
                        "irr": payload["summary_metrics"].get("irr_pct", 0),
                        "payback": payload["summary_metrics"].get("payback_period"),
                        "years": 5,
                        "total_benefit": payload["annual_total"]["total_estimated_benefit"],
                        "total_investment": payload["annual_total"]["total_investment"],
                        "total_cumulative": payload["annual_total"]["cumulative_net_benefit"],
                        "selection": [m["module_name"] for m in payload.get("module_results", []) if m.get("selected")],
                        "module_results": {m["module_name"]: m for m in payload.get("module_results", [])}
                    }
                    # Mostrar reporte inmediatamente
                    st.session_state.excel_results = report_data
                    show_executive_report(report_data)

                except Exception as exc:
                    st.error(f"No fue posible calcular el modelo: {exc}")

        if "excel_results" in st.session_state:
            st.write("---")
            if st.button("📊 Ver Reporte Ejecutivo (Excel)", use_container_width=True):
                show_executive_report(st.session_state.excel_results)
            st.info("Cálculo completado. El detalle completo está disponible en el Reporte Ejecutivo pop-up.")

    if mode == "Entrada manual":
        st.caption("Completa los datos clave y calcula ROI automáticamente sin usar Excel.")
        
        st.subheader("1. Horizonte de análisis")
        st.info("Define el tiempo de retorno. Cambiar este valor actualizará automáticamente el configurador de inversiones.")
        adoption_years = st.radio("Horizonte de análisis (años)", [3, 5], index=1, key="manual_horizon", horizontal=True)

        with st.form("manual_inputs"):
            st.subheader("1. Datos base")
            client_name = st.text_input("Cliente", "Cliente Demo")
            retailer_type = st.selectbox("Tipo de retailer", ["Retail general", "Supermercado", "Moda", "Electrónica"])

            col1, col2, col3 = st.columns(3)
            with col1:
                net_revenue = st.number_input("Ingresos brutos anuales (USD)", min_value=0.0, value=100_000_000.0, step=1_000.0)
                cogs_pct = st.number_input("% COGS sobre ingresos", min_value=0.0, max_value=100.0, value=55.0, step=0.1)
                sga_pct = st.number_input("% SGA sobre ingresos", min_value=0.0, max_value=100.0, value=15.0, step=0.1)
            with col2:
                tax_rate = st.number_input("% tasa de impuestos", min_value=0.0, max_value=100.0, value=30.0, step=0.1)
                inventory = st.number_input("Inventario total (USD)", min_value=0.0, value=25_000_000.0, step=1000.0)
                carrying_cost = st.number_input("% costo de mantenimiento de inventario", min_value=0.0, max_value=100.0, value=20.0, step=0.1)
            with col3:
                growth_rate = st.number_input("% crecimiento de ingresos anual", min_value=-100.0, max_value=100.0, value=3.0, step=0.1)
                discount_rate = st.number_input("% tasa de descuento NPV", min_value=0.0, max_value=100.0, value=10.0, step=0.1)
                # adoption_years movido arriba del form para reactividad inmediata

            st.subheader("2. Selección de módulos Oracle")
            module_options = st.session_state.module_options
            module_selected = st.multiselect("Selecciona módulos Oracle", module_options, default=module_options[:2])

            module_profiles = st.session_state.module_profiles
            benefit_params = st.session_state.benefit_params

            module_benefits = {}
            if module_selected:
                st.write("Usando parámetros configurados.")
                scenario = st.radio("Escenario", ["Conservative", "Base", "Aggressive"], index=1)
                for module in module_selected:
                    module_benefits[module] = {}
                    for aspect in module_profiles.get(module, []):
                        params = benefit_params.get(module, {}).get(aspect, {"min": 5.0, "max": 15.0})
                        if scenario == "Conservative":
                            pct = params["min"]
                        elif scenario == "Base":
                            pct = (params["min"] + params["max"]) / 2
                        else:  # Aggressive
                            pct = params["max"]
                        module_benefits[module][aspect] = pct
            else:
                st.info("Selecciona al menos un módulo para obtener resultados por módulo.")

            st.subheader("3. Inversiones por año")
            st.info("Utiliza el botón '🛠️ Inversiones Anuales' en la barra lateral para ajustar el cronograma multi-anual.")
            
            # Resumen visual discreto
            total_inv_setup = get_total_manual_investment()
            st.caption(f"**Total de Inversión**: {format_usd(total_inv_setup)}")

            submitted = st.form_submit_button("Calcular ROI manual")

        if submitted:
            def calc_manual_roi():
                cogs = net_revenue * cogs_pct / 100.0
                sga = net_revenue * sga_pct / 100.0
                operating_income = net_revenue - cogs - sga
                base_margin = (operating_income / net_revenue * 100.0) if net_revenue > 0 else 0.0

                years = adoption_years
                yearly_total_benefit = []
                yearly_total_investment = []
                yearly_total_cashflow = []
                yearly_total_cumulative = []
                module_results = {m: {"benefit": [], "investment": [], "cashflow": [], "cumulative": []} for m in module_selected}

                # Calcular % total por aspecto (complementario)
                aspect_totals = {"sales": 0.0, "inventory_reduction": 0.0, "margin": 0.0, "labor": 0.0, "logistics": 0.0}
                for module in module_selected:
                    for aspect, pct in module_benefits.get(module, {}).items():
                        aspect_totals[aspect] += pct

                cumulative_total = 0.0
                for year in range(1, years + 1):
                    rev_year = net_revenue * ((1 + growth_rate / 100.0) ** (year - 1))
                    cogs_year = rev_year * cogs_pct / 100.0
                    sga_year = rev_year * sga_pct / 100.0
                    base_oi_year = rev_year - cogs_year - sga_year

                    inv_data = st.session_state.annual_investments.get(year, {'software': 0, 'impl': 0, 'extra': 0})
                    total_investment = inv_data['software'] + inv_data['impl'] + inv_data['extra']

                    # Beneficios totales por aspecto
                    sales_benefit_total = rev_year * (aspect_totals["sales"] / 100.0)
                    inventory_benefit_total = inventory * (aspect_totals["inventory_reduction"] / 100.0) * (carrying_cost / 100.0)
                    margin_benefit_total = base_oi_year * (aspect_totals["margin"] / 100.0)
                    labor_benefit_total = sga_year * (aspect_totals["labor"] / 100.0)
                    logistics_benefit_total = base_oi_year * (aspect_totals["logistics"] / 100.0)

                    total_benefit = sales_benefit_total + inventory_benefit_total + margin_benefit_total + labor_benefit_total + logistics_benefit_total

                    # Distribuir beneficios por módulo
                    for module in module_selected:
                        benefit_module = 0.0
                        for aspect in module_profiles.get(module, []):
                            pct_module_aspect = module_benefits.get(module, {}).get(aspect, 0.0)
                            if aspect_totals[aspect] > 0:
                                if aspect == "sales":
                                    benefit_module += sales_benefit_total * (pct_module_aspect / aspect_totals[aspect])
                                elif aspect == "inventory_reduction":
                                    benefit_module += inventory_benefit_total * (pct_module_aspect / aspect_totals[aspect])
                                elif aspect == "margin":
                                    benefit_module += margin_benefit_total * (pct_module_aspect / aspect_totals[aspect])
                                elif aspect == "labor":
                                    benefit_module += labor_benefit_total * (pct_module_aspect / aspect_totals[aspect])
                                elif aspect == "logistics":
                                    benefit_module += logistics_benefit_total * (pct_module_aspect / aspect_totals[aspect])

                        investment_module = total_investment / len(module_selected) if module_selected else 0.0
                        cashflow_module = benefit_module - investment_module

                        module_results[module]["benefit"].append(round(benefit_module, 2))
                        module_results[module]["investment"].append(round(investment_module, 2))
                        module_results[module]["cashflow"].append(round(cashflow_module, 2))
                        cumulative_module = (module_results[module]["cumulative"][-1] if module_results[module]["cumulative"] else 0.0) + cashflow_module
                        module_results[module]["cumulative"].append(round(cumulative_module, 2))

                    yearly_total_benefit.append(round(total_benefit, 2))
                    yearly_total_investment.append(round(total_investment, 2))
                    cashflow_total = total_benefit - total_investment
                    yearly_total_cashflow.append(round(cashflow_total, 2))
                    cumulative_total += cashflow_total
                    yearly_total_cumulative.append(round(cumulative_total, 2))

                npv_value = npv(discount_rate / 100.0, yearly_total_cashflow)
                irr_value = irr(yearly_total_cashflow)
                payback = payback_period(yearly_total_cumulative)

                return {
                    "cliente": client_name,
                    "retailer_type": retailer_type,
                    "base_net_revenue": net_revenue,
                    "years": years,
                    "total_benefit": yearly_total_benefit,
                    "total_investment": yearly_total_investment,
                    "total_cashflow": yearly_total_cashflow,
                    "total_cumulative": yearly_total_cumulative,
                    "module_results": module_results,
                    "npv": round(npv_value, 2),
                    "irr": round(irr_value * 100.0, 2) if irr_value is not None else None,
                    "payback": payback,
                    "selection": module_selected,
                    "aspect_pcts": module_benefits
                }

            results = calc_manual_roi()
            
            # Guardar en session_state y mostrar reporte inmediatamente
            st.session_state.manual_results = results
            show_executive_report(results)

        if "manual_results" in st.session_state:
            st.write("---")
            st.success("Análisis completado. El detalle completo está disponible en el Reporte Ejecutivo pop-up.")
            if st.button("🔍 Abrir Reporte Ejecutivo"):
                show_executive_report(st.session_state.manual_results)

