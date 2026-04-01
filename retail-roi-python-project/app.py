from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from retail_roi_model.engine import WorkbookDrivenRetailROIModel, outputs_to_jsonable, npv, irr, payback_period

st.set_page_config(page_title="Retail ROI Model", layout="wide")

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

    mode = st.radio("Selecciona modo de entrada", ["Archivo Excel", "Entrada manual"], index=1)

    if mode == "Archivo Excel":
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

                    st.subheader("Summary Metrics")
                    st.json(payload["summary_metrics"])

                    st.subheader("Annual Total")
                    st.json(payload["annual_total"])

                    st.subheader("Quarterly Total")
                    st.json(payload["quarterly_total"])

                    st.download_button(
                        "Descargar JSON",
                        data=json.dumps(payload, indent=2, ensure_ascii=False),
                        file_name="roi_output.json",
                        mime="application/json",
                    )
                except Exception as exc:
                    st.error(f"No fue posible calcular el modelo: {exc}")

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
                }

            results = calc_manual_roi()

            case_tab, detail_tab = st.tabs(["Caso de negocios", "Detalle datos"])

            years = list(range(1, results['years'] + 1))

            with case_tab:
                st.subheader("KPIs ejecutivos")
                st.metric("NPV (USD)", f"{results['npv']:,.0f}")
                st.metric("IRR (%)", f"{results['irr'] if results['irr'] is not None else 'No definido'}")
                st.metric("Payback estimado", results['payback'] or "Más de horizonte")

                st.write("### Selección de módulos")
                st.write(", ".join(results['selection']) if results['selection'] else "Ninguno seleccionado")

                st.write("### Gráficos financieros")
                total_df = pd.DataFrame({
                    "Beneficio": results['total_benefit'],
                    "Inversión": results['total_investment'],
                    "Flujo neto": results['total_cashflow'],
                    "Cumulado": results['total_cumulative'],
                }, index=years)

                st.line_chart(total_df[['Beneficio', 'Inversión', 'Flujo neto']])
                st.bar_chart(total_df[['Cumulado']])

                st.write("### Beneficios por módulo")
                if results['selection']:
                    mod_df = pd.DataFrame({
                        m: results['module_results'][m]['benefit'] for m in results['selection']
                    }, index=years)
                    st.line_chart(mod_df)

                    for mod in results['selection']:
                        st.write(f"**{mod}** - Beneficio total: {sum(results['module_results'][mod]['benefit']):,.0f} USD, Cashflow total: {sum(results['module_results'][mod]['cashflow']):,.0f} USD")
                else:
                    st.write("No hay módulos seleccionados.")

                try:
                    import io
                    import matplotlib.pyplot as plt
                    from pptx import Presentation
                    from pptx.util import Inches

                    fig, ax = plt.subplots(figsize=(10, 4))
                    ax.plot(years, total_df['Beneficio'], marker='o', label='Beneficio')
                    ax.plot(years, total_df['Inversión'], marker='o', label='Inversión')
                    ax.plot(years, total_df['Flujo neto'], marker='o', label='Flujo neto')
                    ax.set_title('KPIs Totales por Año')
                    ax.set_xlabel('Año')
                    ax.set_ylabel('USD')
                    ax.legend()
                    ax.grid(True)

                    img_buf = io.BytesIO()
                    fig.savefig(img_buf, format='png', bbox_inches='tight')
                    plt.close(fig)
                    img_buf.seek(0)

                    prs = Presentation()
                    slide = prs.slides.add_slide(prs.slide_layouts[5])
                    slide.shapes.title.text = 'Caso de negocios - Retail ROI'

                    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(9), Inches(2.5))
                    tf = textbox.text_frame
                    tf.text = f"Cliente: {results['cliente']} | Retailer: {results['retailer_type']} | Módulos: {', '.join(results['selection'])}"
                    p = tf.add_paragraph(); p.text = f"NPV: {results['npv']:,.0f} USD"
                    p = tf.add_paragraph(); p.text = f"IRR: {results['irr'] if results['irr'] is not None else 'No definido'} %"
                    p = tf.add_paragraph(); p.text = f"Payback: {results['payback']}"

                    slide.shapes.add_picture(img_buf, Inches(0.5), Inches(3.5), width=Inches(9))

                    pptx_buf = io.BytesIO()
                    prs.save(pptx_buf)
                    pptx_buf.seek(0)

                    st.download_button(
                        "Exportar Caso de negocios a PPTX",
                        data=pptx_buf,
                        file_name="caso_negocios_roi.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except ImportError:
                    st.warning("Instala python-pptx y matplotlib para habilitar exportación PPTX")

            with detail_tab:
                st.subheader("Detalle anual")
                detail_df = pd.DataFrame({
                    'Beneficio': results['total_benefit'],
                    'Inversión': results['total_investment'],
                    'Flujo neto': results['total_cashflow'],
                    'Cumulado': results['total_cumulative']
                }, index=years)
                st.table(detail_df)

            st.subheader("Resultado JSON")
            st.json(results)

            st.download_button(
                "Descargar resultado JSON",
                data=json.dumps(results, indent=2, ensure_ascii=False),
                file_name="roi_manual_output.json",
                mime="application/json",
            )
