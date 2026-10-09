from datetime import datetime, date, time
import csv
import json
from io import StringIO
from email.message import EmailMessage

import streamlit as st

from core import process_files, build_closed_email
from feedback_store import load_feedbacks, save_feedback, delete_feedback
from sms_store import load_sms_status, mark_sms_sent, mark_sms_pending

st.set_page_config(page_title="Control Rechazos FME", page_icon="🛠️", layout="wide")
st.title("Control de Rechazos FME · V7")
st.caption("Rechazos, coordinadores, WhatsApp FME, feedback OWS, filtros operativos y correos Closed listos para Outlook.")

MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

TO_RECIPIENTS = [
    "Griselda Dolores Castillo Bacilio <griselda.dolores.castillo@eysglobal.co>",
    "Cristian Camilo Jimenez Pena <camilo.jimenez@huawei.com>",
]
CC_RECIPIENTS = [
    "sebastian.rolong_costa@manpowercolombia.com",
    "ali.villa_costa@manpowercolombia.com",
]


def build_eml(to_recipients, cc_recipients, subject, body):
    msg = EmailMessage()
    msg["To"] = ", ".join(to_recipients)
    if cc_recipients:
        msg["Cc"] = ", ".join(cc_recipients)
    msg["Subject"] = subject
    msg.set_content(body)
    return msg.as_bytes()


def has_feedback(row):
    return row.get("Tiene feedback OWS") == "Sí"


with st.sidebar:
    st.header("Periodo de auditoría")
    year = st.number_input("Año", 2024, 2100, datetime.now().year, 1)
    month = st.selectbox("Mes", range(1, 13), index=datetime.now().month - 1, format_func=lambda x: MONTHS[x-1])
    st.info("La fecha del rechazo siempre se toma de **Audit time**. El feedback OWS lo registras tú desde el prototipo.")

c1, c2 = st.columns(2)
with c1:
    base = st.file_uploader("1. RECHAZOS Maria.xlsx", type=["xlsx"], key="base")
    coordinators = st.file_uploader("3. DATA 2026 - COORDINADORES.xlsx", type=["xlsx"], key="coords")
with c2:
    report = st.file_uploader("2. WO Result Query", type=["xlsx"], key="report")
    cuadrillas = st.file_uploader("4. CUADRILLAS ACT. NUEVO PROY. 2026.xlsx", type=["xlsx"], key="cuadrillas")

if not (base and report and coordinators and cuadrillas):
    st.stop()

feedbacks = load_feedbacks()

try:
    rows, messages, output_bytes, site_directory, fme_directory = process_files(
        base.getvalue(), report.getvalue(), coordinators.getvalue(), cuadrillas.getvalue(),
        int(year), int(month), feedbacks=feedbacks
    )
except Exception as exc:
    st.error(f"No se pudo procesar la información: {exc}")
    st.stop()

new_rows = [r for r in rows if r["Es nuevo"] == "Sí"]
closed_rows = [r for r in rows if r["Es Closed"]]
open_new_rows = [r for r in new_rows if not r["Es Closed"]]
without_feedback = [r for r in rows if not has_feedback(r)]
with_phone = [r for r in rows if r.get("Celular FME")]

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Rechazos del periodo", len(rows))
m2.metric("Nuevos", len(new_rows))
m3.metric("Sin feedback OWS", len(without_feedback))
m4.metric("Para notificar FME", len(open_new_rows))
m5.metric("Closed", len(closed_rows))
m6.metric("FME con celular", len(with_phone))

if not rows:
    st.success("No hay rechazos para el mes y año seleccionados.")
    st.stop()

tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Vista coordinadores", "📋 Rechazos", "📲 Mensajes FME", "📝 Feedback OWS", "🔒 Closed / Exclusiones", "📥 Exportar"
])

with tab0:
    st.subheader("Detalle de actividades")
    st.caption("La vista Sin feedback muestra las actividades en las que todavía no registraste Fecha de feedback + Estado + Feedback de OWS.")

    all_coords = sorted({r.get("Coordinador", "Por validar") for r in rows})
    all_fmes = sorted({r.get("FME ID", "") for r in rows if r.get("FME ID")})
    all_sites = sorted({r.get("Site ID", "") for r in rows if r.get("Site ID")})
    all_statuses = sorted({r.get("Task status", "") for r in rows if r.get("Task status")})
    all_ows_statuses = sorted({r.get("Estado OWS", "") for r in rows if r.get("Estado OWS")})

    f1, f2, f3 = st.columns(3)
    coord_filter = f1.selectbox("Filtrar coordinador", ["Todos"] + all_coords, key="dash_coord")
    fme_filter = f2.multiselect("FME", all_fmes, key="dash_fme")
    site_filter = f3.multiselect("Sitio", all_sites, key="dash_site")

    f4, f5, f6 = st.columns(3)
    task_status_filter = f4.multiselect("Estado de actividad", all_statuses, key="dash_task_status")
    ows_status_filter = f5.multiselect("Estado OWS", all_ows_statuses, key="dash_ows_status")
    scope_filter = f6.selectbox(
        "Vista",
        ["Sin feedback", "Con feedback", "Todos los rechazos", "Solo nuevos", "Solo Closed", "Solo abiertos"],
        key="dash_scope"
    )

    def matches(r, ignore_coord=False):
        if not ignore_coord and coord_filter != "Todos" and r.get("Coordinador") != coord_filter:
            return False
        if fme_filter and r.get("FME ID") not in fme_filter:
            return False
        if site_filter and r.get("Site ID") not in site_filter:
            return False
        if task_status_filter and r.get("Task status") not in task_status_filter:
            return False
        if ows_status_filter and r.get("Estado OWS") not in ows_status_filter:
            return False
        if scope_filter == "Sin feedback" and has_feedback(r):
            return False
        if scope_filter == "Con feedback" and not has_feedback(r):
            return False
        if scope_filter == "Solo nuevos" and r.get("Es nuevo") != "Sí":
            return False
        if scope_filter == "Solo Closed" and not r.get("Es Closed"):
            return False
        if scope_filter == "Solo abiertos" and r.get("Es Closed"):
            return False
        return True

    filtered_rows = [r for r in rows if matches(r)]
    left, right = st.columns([2.3, 1])
    with left:
        detail = []
        for r in filtered_rows:
            detail.append({
                "COORDINADOR": r.get("Coordinador", ""),
                "FME": r.get("FME ID", ""),
                "CELULAR": r.get("Celular original", "") or r.get("Celular FME", ""),
                "SITIO": r.get("Site ID", ""),
                "TASK ID": r.get("Task ID", ""),
                "FECHA DE FEEDBACK": r.get("Fecha de feedback", ""),
                "ESTADO": r.get("Estado OWS", "") or "Sin feedback",
                "FEEDBACK": r.get("Feedback OWS", "") or "Sin feedback",
            })
        st.dataframe(detail, use_container_width=True, hide_index=True, height=470)
        st.caption(f"Mostrando {len(filtered_rows)} actividad(es).")

    with right:
        counts = {}
        for r in rows:
            if matches(r, ignore_coord=True):
                c = r.get("Coordinador", "Por validar")
                counts[c] = counts.get(c, 0) + 1
        summary = [{"Coordinador reporte": k, "Actividades": v} for k, v in sorted(counts.items())]
        st.markdown("#### Resumen dinámico")
        st.dataframe(summary, use_container_width=True, hide_index=True, height=360)
        st.metric("Total general", sum(counts.values()))

with tab1:
    st.subheader("Todos los rechazos del periodo")
    display = []
    for r in rows:
        display.append({
            "Nuevo": r["Es nuevo"],
            "Task status": r["Task status"],
            "Coordinador": r["Coordinador"],
            "Task ID": r["Task ID"],
            "Site ID": r["Site ID"],
            "FME": r["FME ID"],
            "Celular": r.get("Celular original", "") or r.get("Celular FME", ""),
            "Audit time": r["Audit time"].strftime("%d/%m/%Y %H:%M:%S") if r["Audit time"] else "",
            "Fecha feedback": r.get("Fecha de feedback", ""),
            "Estado OWS": r.get("Estado OWS", ""),
            "Feedback OWS": r.get("Feedback OWS", ""),
            "Audit Remark": r["Audit remark"],
        })
    st.dataframe(display, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Mensajes para FME")
    st.caption("Filtra por coordinador, abre WhatsApp con el mensaje precargado y registra el estado del envío.")
    sms_status = load_sms_status()

    def message_contact_state(item):
        keys = item.get("Claves contacto", [])
        if keys and all(sms_status.get(str(key), {}).get("enviado") for key in keys):
            return "Enviado"
        if keys and any(sms_status.get(str(key), {}).get("enviado") for key in keys):
            return "Parcial"
        return "Pendiente"

    msg_coords = sorted({m.get("Coordinador", "Por validar") for m in messages})
    mf1, mf2 = st.columns(2)
    msg_coord_filter = mf1.selectbox("Filtrar por coordinador", ["Todos"] + msg_coords, key="msg_coord_filter")
    msg_state_filter = mf2.selectbox("Estado del mensaje", ["Todos", "Pendiente", "Enviado", "Parcial"], key="msg_state_filter")
    filtered_messages = []
    for item in messages:
        state = message_contact_state(item)
        if msg_coord_filter != "Todos" and item.get("Coordinador") != msg_coord_filter:
            continue
        if msg_state_filter != "Todos" and state != msg_state_filter:
            continue
        item = dict(item)
        item["Estado contacto"] = state
        filtered_messages.append(item)

    sm1, sm2, sm3 = st.columns(3)
    sm1.metric("Mensajes visibles", len(filtered_messages))
    sm2.metric("Pendientes", sum(message_contact_state(m) == "Pendiente" for m in messages))
    sm3.metric("Enviados", sum(message_contact_state(m) == "Enviado" for m in messages))
    if not messages:
        st.info("No hay actividades nuevas y abiertas que requieran mensaje al FME.")
    elif not filtered_messages:
        st.info("No hay mensajes que coincidan con los filtros seleccionados.")
    for i, item in enumerate(filtered_messages):
        state = item.get("Estado contacto", "Pendiente")
        icon = "✅" if state == "Enviado" else ("🟡" if state == "Parcial" else "⏳")
        with st.expander(f"{icon} {item['FME']} · {item['Cantidad rechazos']} rechazo(s) · {item['Coordinador']} · {state}"):
            c1, c2, c3, c4 = st.columns([1.25, 1.15, 0.9, 1.05])
            c1.text_input("Celular FME", item.get("Celular visible", "") or "No encontrado", disabled=True, key=f"phone-{i}-{item['FME']}")
            c2.text_input("Origen celular", item.get("Origen celular", ""), disabled=True, key=f"phone-source-{i}-{item['FME']}")
            c3.text_input("Estado", state, disabled=True, key=f"sms-state-{i}-{item['FME']}")
            if item.get("WhatsApp URL"):
                c4.link_button("📲 Abrir WhatsApp", item["WhatsApp URL"], use_container_width=True)
            else:
                c4.warning("Sin celular")
            st.text_area("Mensaje listo para copiar", item["Mensaje"], height=270, key=f"fme-{i}-{item['FME']}")
            b1, b2, b3 = st.columns([1.1, 1.1, 2])
            if state != "Enviado":
                if b1.button("✅ Marcar enviado", key=f"sent-{i}-{item['FME']}", use_container_width=True):
                    mark_sms_sent(item.get("Claves contacto", []))
                    st.success("Mensaje marcado como enviado.")
                    st.rerun()
            else:
                b1.success("Mensaje enviado")
                dates = [sms_status.get(str(key), {}).get("fecha_envio", "") for key in item.get("Claves contacto", [])]
                last_date = max([value for value in dates if value], default="")
                if last_date:
                    b3.caption(f"Registrado: {last_date.replace('T', ' ')}")
            if state in {"Enviado", "Parcial"}:
                if b2.button("↩️ Marcar pendiente", key=f"pending-{i}-{item['FME']}", use_container_width=True):
                    mark_sms_pending(item.get("Claves contacto", []))
                    st.rerun()

with tab3:
    st.subheader("Registrar feedback realizado en OWS")
    st.caption("Después de dejar el comentario en OWS, registra aquí la fecha, el estado y exactamente el feedback que dejaste. Esto alimenta la vista Sin feedback / Con feedback.")
    if st.session_state.pop("feedback_saved_message", None):
        st.success("✅ Feedback guardado exitosamente.")

    options = {
        f"{r['Task ID']} · {r['Site ID']} · {r['FME ID']} · {r['Coordinador']}": idx
        for idx, r in enumerate(rows)
    }
    selected = st.selectbox("Selecciona la actividad", list(options.keys()), key="feedback_activity")
    row = rows[options[selected]]
    key = row.get("Assignment ID") or row.get("Task ID")

    info1, info2, info3, info4 = st.columns(4)
    info1.text_input("Task ID", row.get("Task ID", ""), disabled=True)
    info2.text_input("Site ID", row.get("Site ID", ""), disabled=True)
    info3.text_input("FME", row.get("FME ID", ""), disabled=True)
    info4.text_input("Coordinador", row.get("Coordinador", ""), disabled=True)

    existing_date = row.get("Fecha de feedback", "")
    try:
        parsed_date = datetime.strptime(existing_date, "%Y-%m-%d").date() if existing_date else date.today()
    except ValueError:
        parsed_date = date.today()
    status_options = ["Pendiente", "Aprobado", "Rechazado", "Otro"]
    existing_status = row.get("Estado OWS", "")
    default_status_idx = status_options.index(existing_status) if existing_status in status_options else 0

    with st.form("feedback_form"):
        fc1, fc2 = st.columns(2)
        feedback_date = fc1.date_input("Fecha de feedback", value=parsed_date)
        feedback_status = fc2.selectbox("Estado", status_options, index=default_status_idx)
        feedback_text = st.text_area("Feedback / comentario dejado en OWS", value=row.get("Feedback OWS", ""), height=180, placeholder="Ej.: Se realizan correcciones sugeridas, para aprobación")
        submitted = st.form_submit_button("💾 Guardar feedback OWS", use_container_width=True)
    if submitted:
        if not feedback_text.strip():
            st.error("Debes escribir el comentario de feedback que dejaste en OWS.")
        else:
            save_feedback(
                key,
                feedback_date.isoformat(),
                feedback_status,
                feedback_text.strip(),
                task_id=row.get("Task ID", ""),
                assignment_id=row.get("Assignment ID", ""),
                site_id=row.get("Site ID", ""),
                fme_id=row.get("FME ID", ""),
                coordinador=row.get("Coordinador", ""),
            )
            st.session_state["feedback_saved_message"] = True
            st.rerun()

    if has_feedback(row):
        if st.button("🗑️ Eliminar registro de feedback", key=f"delete-{key}"):
            delete_feedback(key)
            st.rerun()

with tab4:
    st.subheader("Actividades Closed")
    st.caption("Audit time se llena automáticamente. Tú consultas OWS y agregas únicamente la fecha/hora de cierre.")
    if not closed_rows:
        st.success("No hay actividades Closed rechazadas en este periodo.")
    else:
        options = {f"{r['Task ID']} · {r['Site ID']} · {r['Coordinador']}": i for i, r in enumerate(closed_rows)}
        selected = st.selectbox("Selecciona una actividad", list(options.keys()), key="closed_activity")
        row = closed_rows[options[selected]]

        a, b, c = st.columns(3)
        a.text_input("Task ID", row["Task ID"], disabled=True)
        b.text_input("Site ID", row["Site ID"], disabled=True)
        c.text_input("Coordinador", row["Coordinador"], disabled=True)
        st.text_area("Audit Remark", row["Audit remark"], height=220, disabled=True)
        st.text_area("Propósito generado según Audit Remark", row["Proposito"], height=170, disabled=True)

        dc1, dc2 = st.columns(2)
        close_day = dc1.date_input("Fecha de cierre OWS", value=date.today(), key=f"date-{row['Assignment ID']}")
        close_time = dc2.time_input("Hora de cierre OWS", value=time(0, 0, 0), step=1, key=f"time-{row['Assignment ID']}")
        close_dt = datetime.combine(close_day, close_time)

        email = build_closed_email(row, close_dt.strftime("%Y-%m-%d %H:%M:%S"))
        st.markdown("#### Correo generado")
        st.text_area("Para", "; ".join(TO_RECIPIENTS), height=85, disabled=True, key=f"to-{row['Assignment ID']}")
        cc_value = st.text_area("CC", "; ".join(CC_RECIPIENTS), height=85, key=f"cc-{row['Assignment ID']}")
        subject = st.text_input("Asunto", email["Asunto"], key=f"subject-{row['Assignment ID']}")
        body = st.text_area("Cuerpo", email["Cuerpo"], height=520, key=f"body-{row['Assignment ID']}")

        cc_list = [x.strip() for x in cc_value.split(";") if x.strip()]
        eml_bytes = build_eml(TO_RECIPIENTS, cc_list, subject, body)
        st.download_button(
            "📧 Descargar correo listo para Outlook (.eml)",
            data=eml_bytes,
            file_name=f"{row['Task ID'] or row['Assignment ID']}_closed.eml",
            mime="message/rfc822",
            use_container_width=True,
        )

with tab5:
    month_label = MONTHS[month-1].lower()
    st.download_button(
        "Descargar RECHAZOS Maria actualizado",
        data=output_bytes,
        file_name=f"RECHAZOS_Maria_actualizado_{month_label}_{year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    text = [["Coordinador", "FME", "Celular FME", "Cantidad rechazos", "Sitios", "Mensaje", "Estado contacto"]]
    for x in messages:
        text.append([x["Coordinador"], x["FME"], x.get("Celular visible", ""), x["Cantidad rechazos"], x["Sitios"], x["Mensaje"], x["Estado contacto"]])
    string_io = StringIO()
    csv.writer(string_io).writerows(text)
    st.download_button(
        "Descargar mensajes FME en CSV",
        data=string_io.getvalue().encode("utf-8-sig"),
        file_name=f"mensajes_fme_{month_label}_{year}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    feedback_export = []
    current_feedbacks = load_feedbacks()
    for row in rows:
        feedback_key = str(row.get("Assignment ID") or row.get("Task ID") or "")
        feedback = current_feedbacks.get(feedback_key) or current_feedbacks.get(str(row.get("Task ID") or "")) or {}
        if not feedback:
            continue
        feedback_export.append({
            "assignment_id": row.get("Assignment ID", ""),
            "site_id": row.get("Site ID", ""),
            "fme_id": row.get("FME ID", ""),
            "coordinador": row.get("Coordinador", ""),
            "fecha_feedback": feedback.get("fecha", row.get("Fecha de feedback", "")),
            "estado": feedback.get("estado", row.get("Estado OWS", "")),
            "feedback": feedback.get("feedback", row.get("Feedback OWS", "")),
        })
    st.download_button(
        "Descargar feedback OWS en JSON",
        data=json.dumps(feedback_export, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name=f"feedback_ows_{month_label}_{year}.json",
        mime="application/json",
        use_container_width=True,
    )

st.divider()
st.caption("MVP V7 · Coordinador por DATA COORDINADORES · Celular/WhatsApp por CUADRILLAS · Feedback OWS persistente · Closed y correos Outlook.")
