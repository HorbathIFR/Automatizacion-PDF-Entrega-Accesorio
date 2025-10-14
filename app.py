import streamlit as st
import pandas as pd
from datetime import datetime
from weasyprint import HTML

# =================== FUNCIONES COMUNES ===================

def cargar_html(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def llenar_html(template, datos):
    for clave, valor in datos.items():
        template = template.replace(f"{{{{{clave}}}}}", str(valor))
    return template

# ✅ Función para limpiar valores vacíos
def limpiar_valor(valor):
    """Convierte valores None, NaN, o vacíos a 'N/A'"""
    if pd.isna(valor) or valor == "" or valor is None:
        return "N/A"
    return str(valor).strip()

# =================== SIDEBAR ===================

st.sidebar.title("📋 Formularios")
seccion = st.sidebar.selectbox("Selecciona una sección:", [
    "📥 Recepción de equipos",
    "📤 Entrega de equipos"
])

# =================== RECEPCIÓN DE EQUIPOS ===================

if seccion == "📥 Recepción de equipos":
    st.title("📄 Acta de Recepción de Equipos y Accesorios")

    tipo_acta = st.radio(
        "Tipo de acta de recepción:",
        ["Solo Equipo", "Solo Accesorios", "Equipo + Accesorios"]
    )

    # Sección EQUIPO
    df_equipos = None
    equipo_data = None
    serial_equipo = None
    
    if tipo_acta in ["Solo Equipo", "Equipo + Accesorios"]:
        st.subheader("🖥️ Datos del Equipo")
        uploaded_equipos = st.file_uploader(
            "📁 Sube el archivo Excel de equipos", 
            type=["xlsx"], 
            key="equipos_recep_excel"
        )
        
        if uploaded_equipos:
            df_equipos = pd.read_excel(uploaded_equipos)
            st.success(f"✅ Archivo de equipos cargado: {len(df_equipos)} registros")
            serial_equipo = st.text_input("🔍 Número de serie del equipo (Título)")

    # Sección ACCESORIOS
    df_accesorios = None
    accesorios_data = []
    
    if tipo_acta in ["Solo Accesorios", "Equipo + Accesorios"]:
        st.subheader("🔌 Datos de Accesorios")
        uploaded_accesorios = st.file_uploader(
            "📁 Sube el archivo Excel de accesorios", 
            type=["xlsx"], 
            key="accesorios_recep_excel"
        )
        
        if st.session_state.get("tipo_acta_recep") != tipo_acta:
            st.session_state.tipo_acta_recep = tipo_acta
            st.session_state.accesorios_recep_list = [{"titulo": ""}]
        
        if uploaded_accesorios:
            df_accesorios = pd.read_excel(uploaded_accesorios)
            st.success(f"✅ Archivo de accesorios cargado: {len(df_accesorios)} registros")
            
            if 'accesorios_recep_list' not in st.session_state:
                st.session_state.accesorios_recep_list = [{"titulo": ""}]
            
        # Asegurar que la lista de accesorios en sesión sea válida
        if not isinstance(st.session_state.get("accesorios_recep_list", []), list):
            st.session_state.accesorios_recep_list = [{"titulo": ""}]
        else:
            st.session_state.accesorios_recep_list = [
                acc if isinstance(acc, dict) and "titulo" in acc else {"titulo": ""}
                for acc in st.session_state.accesorios_recep_list
            ]

        # Mostrar lista dinámica de accesorios
        for i, acc in enumerate(st.session_state.accesorios_recep_list):
            col1, col2 = st.columns([4, 1])
            with col1:
                st.session_state.accesorios_recep_list[i]["titulo"] = st.text_input(
                    f"🔍 Título del accesorio #{i+1}",
                    value=acc.get("titulo", ""),
                    key=f"acc_recep_{i}"
                )
            with col2:
                if st.button("❌", key=f"remove_recep_{i}"):
                    st.session_state.accesorios_recep_list.pop(i)
                    st.rerun()

        # Botón para agregar más accesorios
        if st.button("➕ Agregar otro accesorio", key="add_recep"):
            st.session_state.accesorios_recep_list.append({"titulo": ""})
            st.rerun()

    # ✅ NUEVO: Motivo de entrega
    st.subheader("📋 Motivo de entrega")
    col1, col2 = st.columns(2)
    with col1:
        motivo_desvinculacion = st.checkbox("Desvinculación", key="recep_desvinculacion")
        motivo_falla_recep = st.checkbox("Cambio por falla o daño", key="recep_falla")
    with col2:
        motivo_renovacion_recep = st.checkbox("Renovación de equipo", key="recep_renovacion")
        motivo_otro_recep = st.checkbox("Otro", key="recep_otro")

    motivo_otro_texto_recep = ""
    if motivo_otro_recep:
        motivo_otro_texto_recep = st.text_input("Especificar otro motivo:", key="otro_motivo_recep")

    # 🔧 Estado del equipo (se diligencia de forma presencial)
    estado_funcional = False
    estado_fallas = False
    estado_danado = False
    estado_fallas_descripcion = ""

    # Datos de las firmas
    st.subheader("✍️ Nombres para las firmas")
    col1, col2 = st.columns(2)
    
    with col1:
        nombre_quien_entrega = st.text_input("Nombre de quien ENTREGA")
    
    with col2:
        nombre_quien_recibe = st.text_input("Nombre de quien RECIBE", value="Jonathan David Santos Arrieta")
        cargo_quien_recibe = st.text_input("Cargo de quien RECIBE", key="cargo_recibe_recepcion")

    # GENERAR PDF
    if st.button("📥 Generar PDF de Recepción"):
        try:
            # Validaciones
            if tipo_acta in ["Solo Equipo", "Equipo + Accesorios"] and df_equipos is None:
                st.error("⚠️ Por favor sube el archivo Excel de equipos.")
                st.stop()
            
            if tipo_acta in ["Solo Accesorios", "Equipo + Accesorios"] and df_accesorios is None:
                st.error("⚠️ Por favor sube el archivo Excel de accesorios.")
                st.stop()

            # Procesar EQUIPO
            if tipo_acta in ["Solo Equipo", "Equipo + Accesorios"] and df_equipos is not None:
                if not serial_equipo or not serial_equipo.strip():
                    st.error("⚠️ Por favor ingresa el número de serie del equipo.")
                    st.stop()
                
                # ✅ Limpiar nombres de columnas
                df_equipos.columns = df_equipos.columns.str.strip()
                
                serial_input = serial_equipo.strip().upper()
                df_equipos["Título"] = df_equipos["Título"].astype(str).str.strip().str.upper()
                equipo = df_equipos[df_equipos["Título"] == serial_input]
                
                if equipo.empty:
                    st.error("⚠️ No se encontró el equipo con ese número de serie.")
                    st.stop()
                
                row = equipo.iloc[0]
                equipo_data = {
                    "n_inventario": limpiar_valor(row["Título"]),
                    "dispositivo": limpiar_valor(row.get("Tipo de activo")),
                    "marca": limpiar_valor(row.get("Fabricante")),
                    "modelo": limpiar_valor(row.get("Modelo")),
                    "serial": limpiar_valor(row.get("Número de serie")),
                    "memoria": limpiar_valor(row.get("RAM")),
                    "procesador": limpiar_valor(row.get("Modelo de procesador")),
                    "almacenamiento": limpiar_valor(row.get("Capacidad"))
                }

            # Procesar ACCESORIOS
            if tipo_acta in ["Solo Accesorios", "Equipo + Accesorios"] and df_accesorios is not None:
                # ✅ Limpiar nombres de columnas
                df_accesorios.columns = df_accesorios.columns.str.strip()
                df_accesorios["Título"] = df_accesorios["Título"].astype(str).str.strip().str.upper()
                
                for acc in st.session_state.accesorios_recep_list:
                    if acc.get("titulo", "").strip():
                        titulo_acc = acc["titulo"].strip().upper()
                        acc_row = df_accesorios[df_accesorios["Título"] == titulo_acc]
                        
                        if not acc_row.empty:
                            row = acc_row.iloc[0]
                            accesorios_data.append({
                                "tipo": limpiar_valor(row.get("Tipo de activo")),
                                "marca": limpiar_valor(row.get("Fabricante")),
                                "modelo": limpiar_valor(row.get("Modelo")),
                                "serial": limpiar_valor(row.get("Número de serie")),
                                "n_inventario": limpiar_valor(row.get("Título"))
                            })

            # ✅ Construir fila del equipo (solo si existe equipo_data)
            equipo_row = ""
            if equipo_data:
                equipo_row = f"""
                <tr>
                    <td style="width: 13%;">{equipo_data['n_inventario']}</td>
                    <td style="width: 12%;">{equipo_data['dispositivo']}</td>
                    <td style="width: 8%;">{equipo_data['marca']}</td>
                    <td style="width: 15%;">{equipo_data['modelo']}</td>
                    <td style="width: 14%;">{equipo_data['serial']}</td>
                    <td style="width: 11%;">{equipo_data['memoria']}</td>
                    <td style="width: 19%;">{equipo_data['procesador']}</td>
                    <td style="width: 8%;">{equipo_data['almacenamiento']}</td>
                </tr>
                """

            # ✅ Construir filas de accesorios (sin filas vacías extras)
            accesorios_html = ""
            if accesorios_data:
                for acc in accesorios_data:
                    accesorios_html += f"""
                    <tr>
                        <td style="width: 13%;">{acc['n_inventario']}</td>
                        <td style="width: 12%;">{acc['tipo']}</td>
                        <td style="width: 8%;">{acc['marca']}</td>
                        <td style="width: 15%;">{acc['modelo']}</td>
                        <td style="width: 14%;">{acc['serial']}</td>
                        <td style="width: 11%;">N/A</td>
                        <td style="width: 19%;">N/A</td>
                        <td style="width: 8%;">N/A</td>
                    </tr>
                    """
            
            # ✅ Construir checkboxes de motivo
            check_desvinculacion = "☑" if motivo_desvinculacion else "☐"
            check_renovacion_recep = "☑" if motivo_renovacion_recep else "☐"
            check_falla_recep = "☑" if motivo_falla_recep else "☐"
            check_otro_recep = "☑" if motivo_otro_recep else "☐"

            # ✅ Construir checkboxes de estado
            check_funcional = "☑" if estado_funcional else "☐"
            check_con_fallas = "☑" if estado_fallas else "☐"
            check_danado = "☑" if estado_danado else "☐"

            # Construir datos para el HTML
            datos_pdf = {
                "fecha_actual": datetime.now().strftime("%d/%m/%Y"),
                "equipo_row": equipo_row,
                "accesorios_rows": accesorios_html,
                "nombre_quien_entrega": nombre_quien_entrega,
                "nombre_quien_recibe": nombre_quien_recibe,
                "cargo_quien_recibe": cargo_quien_recibe,
                "check_desvinculacion": check_desvinculacion,
                "check_renovacion": check_renovacion_recep,
                "check_falla": check_falla_recep,
                "check_otro": check_otro_recep,
                "motivo_otro_texto": motivo_otro_texto_recep,
                "check_funcional": check_funcional,
                "check_con_fallas": check_con_fallas,
                "check_danado": check_danado,
                "fallas_descripcion": estado_fallas_descripcion
            }

            html_template = cargar_html("recepcion_v3.html")
            html_lleno = llenar_html(html_template, datos_pdf)
            pdf_bytes = HTML(string=html_lleno, base_url=".").write_pdf()

            st.success("✅ PDF generado exitosamente.")
            st.download_button(
                label="⬇️ Descargar PDF",
                data=pdf_bytes,
                file_name=f"Acta_Recepcion_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"❌ Error al generar PDF: {str(e)}")

# =================== ENTREGA DE EQUIPOS ===================

elif seccion == "📤 Entrega de equipos":
    st.title("📦 Acta de Entrega de Equipos y Accesorios")

    tipo_acta = st.radio(
        "Tipo de acta de entrega:",
        ["Solo Equipo", "Solo Accesorios", "Equipo + Accesorios"]
    )

    # Sección EQUIPO
    df_equipos = None
    equipo_data = None
    serial_equipo = None
    
    if tipo_acta in ["Solo Equipo", "Equipo + Accesorios"]:
        st.subheader("🖥️ Datos del Equipo")
        uploaded_equipos = st.file_uploader(
            "📁 Sube el archivo Excel de equipos", 
            type=["xlsx"], 
            key="equipos_entrega_excel"
        )
        
        if uploaded_equipos:
            df_equipos = pd.read_excel(uploaded_equipos)
            st.success(f"✅ Archivo de equipos cargado: {len(df_equipos)} registros")
            serial_equipo = st.text_input("🔍 Número de serie del equipo (Título)")

    # Sección ACCESORIOS
    df_accesorios = None
    accesorios_data = []
    
    if tipo_acta in ["Solo Accesorios", "Equipo + Accesorios"]:
        st.subheader("🔌 Datos de Accesorios")
        uploaded_accesorios = st.file_uploader(
            "📁 Sube el archivo Excel de accesorios", 
            type=["xlsx"], 
            key="accesorios_entrega_excel"
        )
        
        # Validar cambio de tipo de acta
        if st.session_state.get("tipo_acta_entrega") != tipo_acta:
            st.session_state.tipo_acta_entrega = tipo_acta
            st.session_state.accesorios_entrega = [{"titulo": ""}]
        
        if uploaded_accesorios:
            df_accesorios = pd.read_excel(uploaded_accesorios)
            st.success(f"✅ Archivo de accesorios cargado: {len(df_accesorios)} registros")
        
        # Inicializar lista si no existe
        if 'accesorios_entrega' not in st.session_state:
            st.session_state.accesorios_entrega = [{"titulo": ""}]
        
        # Asegurar que la lista sea válida
        if not isinstance(st.session_state.get("accesorios_entrega", []), list):
            st.session_state.accesorios_entrega = [{"titulo": ""}]
        else:
            st.session_state.accesorios_entrega = [
                acc if isinstance(acc, dict) and "titulo" in acc else {"titulo": ""}
                for acc in st.session_state.accesorios_entrega
            ]
        
        # Mostrar lista dinámica de accesorios
        for i, acc in enumerate(st.session_state.accesorios_entrega):
            col1, col2 = st.columns([4, 1])
            with col1:
                st.session_state.accesorios_entrega[i]["titulo"] = st.text_input(
                    f"🔍 Título del accesorio #{i+1}",
                    value=acc.get("titulo", ""),
                    key=f"acc_entrega_{i}"
                )
            with col2:
                if st.button("❌", key=f"remove_entrega_{i}"):
                    st.session_state.accesorios_entrega.pop(i)
                    st.rerun()
        
        # Botón para agregar más accesorios
        if st.button("➕ Agregar otro accesorio", key="add_entrega"):
            st.session_state.accesorios_entrega.append({"titulo": ""})
            st.rerun()

    # Motivo de entrega
    st.subheader("📋 Motivo de entrega")
    col1, col2 = st.columns(2)
    with col1:
        motivo_vinculacion = st.checkbox("Nueva vinculación")
        motivo_falla = st.checkbox("Cambio por falla o daño")
    with col2:
        motivo_renovacion = st.checkbox("Renovación de equipo")
        motivo_otro = st.checkbox("Otro")

    motivo_otro_texto = ""
    if motivo_otro:
        motivo_otro_texto = st.text_input("Especificar otro motivo:", key="otro_motivo_entrega")

    # Datos de las firmas
    st.subheader("✍️ Nombres para las firmas")
    col1, col2 = st.columns(2)
    
    with col1:
        nombre_quien_recibe = st.text_input("Nombre de quien RECIBE")
    
    with col2:
        nombre_quien_entrega = st.text_input("Nombre de quien ENTREGA", value="Jonathan David Santos Arrieta")
        cargo_quien_entrega = st.text_input("Cargo de quien ENTREGA", value="Coordinador de Infraestructura TI", key="cargo_entrega_entrega")

    # GENERAR PDF
    if st.button("📤 Generar PDF de Entrega"):
        try:
            # Validaciones
            if tipo_acta in ["Solo Equipo", "Equipo + Accesorios"] and df_equipos is None:
                st.error("⚠️ Por favor sube el archivo Excel de equipos.")
                st.stop()
            
            if tipo_acta in ["Solo Accesorios", "Equipo + Accesorios"] and df_accesorios is None:
                st.error("⚠️ Por favor sube el archivo Excel de accesorios.")
                st.stop()

            # Procesar EQUIPO
            if tipo_acta in ["Solo Equipo", "Equipo + Accesorios"] and df_equipos is not None:
                if not serial_equipo or not serial_equipo.strip():
                    st.error("⚠️ Por favor ingresa el número de serie del equipo.")
                    st.stop()
                
                # ✅ Limpiar nombres de columnas (elimina espacios extra)
                df_equipos.columns = df_equipos.columns.str.strip()
                
                serial_input = serial_equipo.strip().upper()
                df_equipos["Título"] = df_equipos["Título"].astype(str).str.strip().str.upper()
                equipo = df_equipos[df_equipos["Título"] == serial_input]
                
                if equipo.empty:
                    st.error("⚠️ No se encontró el equipo con ese número de serie.")
                    st.stop()
                
                row = equipo.iloc[0]
                fabricante = limpiar_valor(row.get("Fabricante"))
                modelo = limpiar_valor(row.get("Modelo"))
                
                # Concatenar marca y modelo, manejando N/A
                if fabricante == "N/A" and modelo == "N/A":
                    marca_modelo = "N/A"
                elif fabricante == "N/A":
                    marca_modelo = modelo
                elif modelo == "N/A":
                    marca_modelo = fabricante
                else:
                    marca_modelo = f"{fabricante} {modelo}"
                
                equipo_data = {
                    "tipo_equipo": limpiar_valor(row.get("Tipo de activo")),
                    "marca_modelo": marca_modelo,
                    "serial": limpiar_valor(row.get("Número de serie")),
                    "procesador": limpiar_valor(row.get("Modelo de procesador")),
                    "memoria": limpiar_valor(row.get("RAM")),
                    "almacenamiento": limpiar_valor(row.get("Capacidad")),
                    "inventario": limpiar_valor(row["Título"])
                }

            # Procesar ACCESORIOS
            if tipo_acta in ["Solo Accesorios", "Equipo + Accesorios"] and df_accesorios is not None:
                # ✅ Limpiar nombres de columnas
                df_accesorios.columns = df_accesorios.columns.str.strip()
                df_accesorios["Título"] = df_accesorios["Título"].astype(str).str.strip().str.upper()
                
                for acc in st.session_state.get("accesorios_entrega", []):
                    if acc.get("titulo", "").strip():
                        titulo_acc = acc["titulo"].strip().upper()
                        acc_row = df_accesorios[df_accesorios["Título"] == titulo_acc]
                        
                        if not acc_row.empty:
                            row = acc_row.iloc[0]
                            accesorios_data.append({
                                "tipo": limpiar_valor(row.get("Tipo de activo")),
                                "marca": limpiar_valor(row.get("Fabricante")),
                                "modelo": limpiar_valor(row.get("Modelo")),
                                "serial": limpiar_valor(row.get("Número de serie")),
                                "n_inventario": limpiar_valor(row.get("Título"))
                            })

            # ✅ Construir filas de accesorios dinámicas
            accesorios_html = ""
            if accesorios_data:
                for acc in accesorios_data:
                    accesorios_html += f"""
                    <tr>
                        <td>{acc['tipo']}</td>
                        <td>{acc['marca']}</td>
                        <td>{acc['modelo']}</td>
                        <td>{acc['serial']}</td>
                        <td>{acc['n_inventario']}</td>
                    </tr>
                    """
            else:
                accesorios_html = '<tr><td colspan="5" style="text-align: center; font-style: italic; color: #666;">N/A</td></tr>'

            # Construir checkboxes
            check_vinculacion = "☑" if motivo_vinculacion else "☐"
            check_renovacion = "☑" if motivo_renovacion else "☐"
            check_falla = "☑" if motivo_falla else "☐"
            check_otro = "☑" if motivo_otro else "☐"

            # Construir datos PDF
            datos_pdf = {
                "fecha_actual": datetime.now().strftime("%d/%m/%Y"),
                "accesorios_rows": accesorios_html,
                "nombre_quien_recibe": nombre_quien_recibe,
                "nombre_quien_entrega": nombre_quien_entrega,
                "cargo_quien_entrega": cargo_quien_entrega,
                "check_vinculacion": check_vinculacion,
                "check_renovacion": check_renovacion,
                "check_falla": check_falla,
                "check_otro": check_otro,
                "motivo_otro_texto": motivo_otro_texto
            }

            if equipo_data:
                datos_pdf.update(equipo_data)
            else:
                datos_pdf.update({
                    "tipo_equipo": "N/A",
                    "marca_modelo": "N/A",
                    "serial": "N/A",
                    "procesador": "N/A",
                    "memoria": "N/A",
                    "almacenamiento": "N/A",
                    "inventario": "N/A"
                })

            html_template = cargar_html("entrega_v3.html")
            html_lleno = llenar_html(html_template, datos_pdf)
            pdf_bytes = HTML(string=html_lleno, base_url=".").write_pdf()

            st.success("✅ PDF generado exitosamente.")
            st.download_button(
                label="⬇️ Descargar PDF",
                data=pdf_bytes,
                file_name=f"Acta_Entrega_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"❌ Error al generar PDF: {str(e)}")
