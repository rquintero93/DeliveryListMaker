import warnings
import io
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

def fill_ga_code(row):
    if pd.isna(row["Código GA Programa"]):
        return row["SF GA_y"]
    else:
        return row["Código GA Programa"]

def fill_contact_email(row):
    if pd.isna(row["Contact Email"]):
        return row["Email Address"]
    else:
        return row["Contact Email"]

def fill_first_name(row):
    if pd.isna(row["Contact First Name"]):
        return row["Person Account: First Name"]
    else:
        return row["Contact First Name"]

def fill_last_name(row):
    if pd.isna(row["Contact Last Name"]):
        return row["Person Account: Last Name"]
    else:
        return row["Contact Last Name"]

def fill_id_contacto(row):
    if pd.isna(row["ID contacto"]):
        return row["Account ID"]
    else:
        return row["ID contacto"]

def fill_id_oportunidad(row):
    if pd.isna(row["Id. de la oportunidad"]):
        return row["Online Application ID"]
    else:
        return row["Id. de la oportunidad"]

def fill_language(row):
    if pd.isna(row["Account Language"]):
        return row["Program Language"]
    else:
        return row["Account Language"]

def tipologia_programa(row):
    if pd.isna(row["Account Tipología de programa"]):
        return row["LOB Category"]
    else:
        return row["Account Tipología de programa"]

def fill_stage(row):
    if pd.isna(row["Estado admisión/inscripción"]):
        return row["Stage"]
    else:
        return row["Estado admisión/inscripción"]

def fill_convocatoria(row):
    if pd.isna(row["Inscrito convocatoria"]):
        return row["Interested Course: Batch Start Month And Year"]
    else:
        return row["Inscrito convocatoria"]

def fill_estado_pago(row):
    if pd.isna(row["Estado pagos"]):
        return row["EE Balance Due"]
    else:
        return row["Estado pagos"]

def fill_gender(row):
    if pd.isna(row["Contact Gender"]):
        return row["Gender"]
    else:
        return row["Contact Gender"]

def fill_country(row):
    if pd.isna(row["ISO - Pais"]):
        return row["Country/Region of Residence"]
    else:
        return row["ISO - Pais"]

def fill_empresa(row):
    if pd.isna(row["Empresa"]):
        return row["Company Name.1"]
    else:
        return row["Empresa"]

def fill_cargo(row):
    if pd.isna(row["Contacto: Cargo"]):
        return row["Job Title"]
    else:
        return row["Contacto: Cargo"]

def sector(row):
    if pd.isna(row["Sector Empresa"]):
        return row["Your Industry"]
    else:
        return row["Sector Empresa"]

def fill_area(row):
    if pd.isna(row["Área/ Departamento"]):
        return row["Your Function"]
    else:
        return row["Área/ Departamento"]

def fill_contact_job_title(row):
    if pd.isna(row["Contact Job Title"]):
        return row["Your Function"]
    else:
        return row["Contact Job Title"]

def fill_experience(row):
    if pd.isna(row["Años de experiencia profesional"]):
        return row["Work Experience"]
    else:
        return row["Años de experiencia profesional"]

def process_files(ee,ga):
    ee = pd.read_excel(ee)
    ga = pd.read_excel(ga)
    codigos = pd.read_csv("Codigos de Universidades.csv")

    ga = ga.merge(codigos, left_on="Código GA Programa", right_on="SF GA", how="inner")
    ee = ee.merge(
        codigos, left_on="Interested Course: Programme", right_on="SF EM", how="inner"
    )

    merged = ga.merge(
        ee,
        left_on=["Contact Email", "SF EM"],
        right_on=["Email Address", "Interested Course: Programme"],
        how="outer",
    )

    merged["Código GA Programa"] = merged.apply(fill_ga_code, axis=1)
    merged["Contact Email"] = merged.apply(fill_contact_email, axis=1)
    merged["Contact First Name"] = merged.apply(fill_first_name, axis=1)
    merged["Contact Last Name"] = merged.apply(fill_last_name, axis=1)
    merged["ID contacto"] = merged.apply(fill_id_contacto, axis=1)
    merged["Id. de la oportunidad"] = merged.apply(fill_id_oportunidad, axis=1)
    merged["Account Language"] = merged.apply(fill_language, axis=1)
    merged["Account Tipología de programa"] = merged.apply(tipologia_programa, axis=1)
    merged["Estado admisión/inscripción"] = merged.apply(fill_stage, axis=1)
    merged["Inscrito convocatoria"] = merged.apply(fill_convocatoria, axis=1)
    merged["Estado pagos"] = merged.apply(fill_estado_pago, axis=1)
    merged["Contact Gender"] = merged.apply(fill_gender, axis=1)
    merged["ISO - Pais"] = merged.apply(fill_country, axis=1)
    merged["Empresa"] = merged.apply(fill_empresa, axis=1)
    merged["Contacto: Cargo"] = merged.apply(fill_cargo, axis=1)
    merged["Sector Empresa"] = merged.apply(sector, axis=1)
    merged["Área/ Departamento"] = merged.apply(fill_area, axis=1)
    merged["Contact Job Title"] = merged.apply(fill_contact_job_title, axis=1)
    merged["Años de experiencia profesional"] = merged.apply(fill_experience, axis=1)

    merged = merged.drop_duplicates(subset=["Contact Email", "Código GA Programa"])

    output = merged[
        [
            "Id. de la oportunidad",
            "ID contacto",
            "Contact First Name",
            "Contact Last Name",
            "Contact Email",
            "Código GA Programa",
            "Account Language",
            "Account Tipología de programa",
            "Estado admisión/inscripción",
            "Inscrito convocatoria",
            "Nueva convocatoria",
            "Notas",
            "Estado pagos",
            "Contact Gender",
            "ISO - Pais",
            "Empresa",
            "Contacto: Cargo",
            "Sector Empresa",
            "Área/ Departamento",
            "Contact Job Title",
            "Años de experiencia profesional",
            # "URL oportunidad",
            # "Marca accion",
        ]
    ]

    return output

st.title("Delivery List Maker")

uploaded_file1 = st.file_uploader("reporte Excel SF EE", type=["xlsx"])
uploaded_file2 = st.file_uploader("reporte Excel SF GA", type=["xlsx"])

if uploaded_file1 is not None and uploaded_file2 is not None:
    st.write("Carga completada.")
    result_df = process_files(uploaded_file1, uploaded_file2)
    st.write("Resultado Final:")
    st.dataframe(result_df)

    # Create an in-memory buffer to save the processed file
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
    buffer.seek(0)

    # Download processed file
    st.download_button(
        label="Descargar archivo procesado",
        data=buffer,
        file_name='Lista Procesada.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
