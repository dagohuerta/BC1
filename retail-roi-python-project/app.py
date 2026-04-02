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

# --- Función para mostrar el Reporte Ejecutivo ---
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

    # --- Botón de Exportación Discreto ---
    try:
        # Generar PPTX en memoria
        prs = Presentation()
        # Slide Título
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blind slide
        
        # Fondo degradado simulado con un rectángulo
        from pptx.enum.shapes import MSO_SHAPE
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.2))
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(31, 119, 180) # Azul corporativo
        shape.line.fill.background()

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(8), Inches(1))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"Resumen Ejecutivo: {results['cliente']}"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # KPIs
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(2) # Ajustado para 4 KPIs
        height = Inches(1)
        
        kpis = [
            ("NPV", format_m(results['npv'])), 
            ("IRR", f"{results['irr']}%"), 
            ("Payback", results['payback']),
            ("Total Beneficios", format_m(sum(results['total_benefit'])))
        ]
        
        for label, val in kpis:
            box = slide.shapes.add_textbox(left, top, width, height)
            tf = box.text_frame
            p_label = tf.paragraphs[0]
            p_label.text = label
            p_label.font.size = Pt(11)
            p_label.font.color.rgb = RGBColor(100, 100, 100)
            
            p_val = tf.add_paragraph()
            p_val.text = val
            p_val.font.size = Pt(16)
            p_val.font.bold = True
            p_val.font.color.rgb = RGBColor(31, 119, 180)
            left += Inches(2.2)

        # Gráfico
        plt.style.use('ggplot')
        fig, ax = plt.subplots(figsize=(8, 4))
        years = list(range(1, results['years'] + 1))
        ax.bar(years, results['total_benefit'], color='#2ca02c', alpha=0.6, label='Beneficios')
        ax.bar(years, results['total_investment'], color='#d62728', alpha=0.6, label='Inversiones')
        ax2 = ax.twinx()
        ax2.plot(years, results['total_cumulative'], color='#1f77b4', linewidth=3, marker='o', label='Flujo Cumulado')
        ax.legend(loc='upper left')
        ax.set_title("Trayectoria Financiera")
        
        img_buf = io.BytesIO()
        fig.savefig(img_buf, format='png', bbox_inches='tight')
        img_buf.seek(0)
        plt.close(fig)
        
        slide.shapes.add_picture(img_buf, Inches(0.5), Inches(3.0), width=Inches(8))

        pptx_buf = io.BytesIO()
        prs.save(pptx_buf)
        pptx_buf.seek(0)

        st.write("---")
        c1, c2 = st.columns([4, 1])
        with c2:
            st.download_button(
                "📥 Exportar PPTX",
                data=pptx_buf,
                file_name=f"ROI_Ejecutivo_{results['cliente']}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                width="stretch",
                help="Descargar este reporte en formato PowerPoint profesional"
            )
    except Exception as e:
        st.warning(f"Exportación PPTX no disponible: {e}")

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
        col_add, col_list = st.columns([1, 2])

        with col_add:
            new_module = st.text_input("Nuevo módulo", key="new_module")
            if st.button("Agregar módulo"):
                if new_module and new_module not in st.session_state.module_options:
                    st.session_state.module_options.append(new_module)
                    st.session_state.module_profiles[new_module] = []
                    st.session_state.benefit_params[new_module] = {}
                    st.success(f"Módulo {new_module} agregado")
                    st.rerun()

        with col_list:
            for i, module in enumerate(st.session_state.module_options):
                col_name, col_aspects, col_actions = st.columns([2, 3, 2])
                with col_name:
                    edited_name = st.text_input(f"Nombre {i}", module, key=f"name_{i}")
                    if edited_name != module:
                        # Renombrar
                        idx = st.session_state.module_options.index(module)
                        st.session_state.module_options[idx] = edited_name
                        st.session_state.module_profiles[edited_name] = st.session_state.module_profiles.pop(module)
                        st.session_state.benefit_params[edited_name] = st.session_state.benefit_params.pop(module)
                        st.rerun()

                with col_aspects:
                    current_aspects = st.session_state.module_profiles.get(module, [])
                    selected_aspects = st.multiselect(f"Aspectos {i}", list(aspect_ranges.keys()), default=current_aspects, key=f"aspects_{i}")
                    if selected_aspects != current_aspects:
                        st.session_state.module_profiles[module] = selected_aspects
                        # Actualizar benefit_params
                        for aspect in selected_aspects:
                            if aspect not in st.session_state.benefit_params[module]:
                                min_val, max_val = aspect_ranges[aspect]
                                st.session_state.benefit_params[module][aspect] = {"min": min_val + 1, "max": max_val - 1}
                        # Remover aspectos no seleccionados
                        for aspect in list(st.session_state.benefit_params[module].keys()):
                            if aspect not in selected_aspects:
                                del st.session_state.benefit_params[module][aspect]

                with col_actions:
                    if st.button(f"Eliminar {i}", key=f"del_{i}"):
                        st.session_state.module_options.remove(module)
                        del st.session_state.module_profiles[module]
                        del st.session_state.benefit_params[module]
                        st.success(f"Módulo {module} eliminado")
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
        st.subheader("Gestión de Aspectos")
        
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
                adoption_years = st.radio("Horizonte de análisis (años)", [3, 5], index=1)

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
            software_fee = st.number_input("Licencia software - año 1", min_value=0.0, value=1_200_000.0, step=1000.0)
            implementation_services = st.number_input("Servicios implementación - año 1", min_value=0.0, value=800_000.0, step=1000.0)

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

                    if year == 1:
                        total_investment = software_fee + implementation_services
                    else:
                        total_investment = 0.0

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

