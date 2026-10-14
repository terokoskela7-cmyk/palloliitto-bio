import datetime
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Palloliitto Bio", page_icon="⚽", layout="centered"
)

st.title("⚽ Palloliitto Bio — PalloID & Pelioikeus")
st.write(
    "Valitse testausmenetelmä, syötä pelaajan 8-numeroinen PalloID ja"
    " syntymäaika. Sovellus tarkistaa automaattisesti yli-ikäisyyssäännön."
)


# Funktio Google Sheetsiin tallennusta varten
def save_to_google_sheets(row_data):
  try:
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    # Luetaan suoraan st.secrets-taulukosta
    creds_dict = dict(st.secrets["gcp_service_account"])

    credentials = Credentials.from_service_account_info(
        creds_dict, scopes=scopes
    )
    client = gspread.authorize(credentials)

    sheet = client.open("Palloliitto Bio Tulokset").sheet1
    sheet.append_row(row_data)
    return True, None
  except Exception as e:
    return False, str(e)


test_type = st.radio(
    "Valitse testausmenetelmä:",
    [
        "Mirwald (PHV & Maturity Offset)",
        "Khamis-Roche (Aikuisiän pituusennuste / PAH)",
    ],
)

with st.form("bio_form"):
  st.subheader("1. Pelaajan tunnistetiedot")
  col1, col2 = st.columns(2)
  with col1:
    pallo_id = st.text_input(
        "PalloID (8 numeroa)", max_chars=8, placeholder="Esim. 12345678"
    )
    gender = st.selectbox("Sukupuoli", ["Poika", "Tyttö"])
  with col2:
    dob = st.date_input(
        "Syntymäaika",
        value=datetime.date(2012, 1, 1),
        min_value=datetime.date(2005, 1, 1),
        max_value=datetime.date(2022, 12, 31),
        format="DD.MM.YYYY",
    )
    dom = st.date_input(
        "Mittauspäivä", value=datetime.date.today(), format="DD.MM.YYYY"
    )

  st.subheader("2. Antropometriset mitat")
  col3, col4 = st.columns(2)
  with col3:
    height = st.number_input(
        "Seisomapituus (cm)",
        min_value=100.0,
        max_value=220.0,
        value=150.0,
        step=0.1,
    )
    weight = st.number_input(
        "Paino (kg)", min_value=20.0, max_value=120.0, value=40.0, step=0.1
    )

  with col4:
    if "Mirwald" in test_type:
      sitting_height = st.number_input(
          "Istumapituus (cm)",
          min_value=50.0,
          max_value=120.0,
          value=75.0,
          step=0.1,
      )
      kr_group = 13
    else:
      sitting_height = 0.0
      mum_height = st.number_input(
          "Äidin pituus (cm)",
          min_value=130.0,
          max_value=200.0,
          value=165.0,
          step=0.5,
      )
      dad_height = st.number_input(
          "Isän pituus (cm)",
          min_value=140.0,
          max_value=220.0,
          value=180.0,
          step=0.5,
      )
      kr_group = st.selectbox("Ikäryhmä (Khamis-Roche)", [13, 14, 15])

  submitted = st.form_submit_button("Laske ja tarkista pelioikeus")

if submitted:
  if not pallo_id or len(pallo_id) != 8 or not pallo_id.isdigit():
    st.warning("Syötä kelvollinen 8-numeroinen PalloID.")
  else:
    age = (dom - dob).days / 365.25
    gender_num = 1 if gender == "Poika" else 2
    birth_month = dob.month
    dom_fi = dom.strftime("%d.%m.%Y")

    if "Mirwald" in test_type:
      leg_length = height - sitting_height
      if gender_num == 1:
        offset = (
            -9.236
            + (0.0002708 * leg_length * sitting_height)
            - (0.001663 * age * leg_length)
            + (0.007216 * age * sitting_height)
            + (0.02292 * (weight / height) * 100)
        )
      else:
        offset = (
            -9.376
            + (0.0001882 * leg_length * sitting_height)
            + (0.0022 * age * leg_length)
            + (0.005841 * age * sitting_height)
            - (0.02658 * (weight / height) * 100)
        )

      phv_age = age - offset

      mirwald_limits_boys = [
          14.966667,
          14.883333,
          14.800000,
          14.716667,
          14.633333,
          14.550000,
          14.466667,
          14.383333,
          14.300000,
          14.216667,
          14.133333,
          14.050000,
      ]
      mirwald_limits_girls = [
          13.066667,
          12.983333,
          12.900000,
          12.816667,
          12.733333,
          12.650000,
          12.566667,
          12.483333,
          12.400000,
          12.316667,
          12.233333,
          12.150000,
      ]

      limit_val = (
          mirwald_limits_boys[birth_month - 1]
          if gender_num == 1
          else mirwald_limits_girls[birth_month - 1]
      )
      eligible = phv_age > limit_val
      eligibility_str = "KYLLÄ" if eligible else "EI"

      st.success(f"Mirwald-tulokset laskettu PalloID:lle {pallo_id}:")
      r1, r2, r3 = st.columns(3)
      r1.metric("Kronologinen ikä", f"{age:.2f} v")
      r2.metric("Arvioitu PHV-ikä", f"{phv_age:.2f} v")
      r3.metric(
          "Pelioikeudellinen (Yli-ikäisyyssääntö)",
          eligibility_str,
          delta="Täyttää vaatimuksen" if eligible else "Ei täytä vaatimusta",
          delta_color="normal" if eligible else "inverse",
      )

      row_to_save = [
          dom_fi,
          "Mirwald",
          pallo_id,
          gender,
          round(age, 2),
          round(phv_age, 2),
          eligibility_str,
      ]

    else:
      mid_parent = (mum_height + dad_height) / 2
      if gender_num == 1:
        pah = (
            -26.052
            + (1.23812 * height)
            - (0.48849 * weight)
            + (0.50286 * mid_parent)
        )
      else:
        pah = (
            -20.657
            + (1.24768 * height)
            - (1.08831 * weight)
            + (0.44774 * mid_parent)
        )

      kr_table = {
          13: [
              11.75,
              11.84,
              11.93,
              12.02,
              12.11,
              12.20,
              12.30,
              12.39,
              12.48,
              12.57,
              12.66,
              12.75,
          ],
          14: [
              12.75,
              12.84,
              12.93,
              13.02,
              13.11,
              13.20,
              13.30,
              13.39,
              13.48,
              13.57,
              13.66,
              13.75,
          ],
          15: [
              13.75,
              13.84,
              13.93,
              14.02,
              14.11,
              14.20,
              14.30,
              14.39,
              14.48,
              14.57,
              14.66,
              14.75,
          ],
      }

      limit_val_kr = kr_table[kr_group][birth_month - 1]
      bio_age_est = age - 0.4 if pah > 175 else age
      eligible_kr = bio_age_est < limit_val_kr
      eligibility_kr_str = "KYLLÄ" if eligible_kr else "EI"

      st.success(f"Khamis-Roche-tulokset laskettu PalloID:lle {pallo_id}:")
      r1, r2, r3 = st.columns(3)
      r1.metric("Kronologinen ikä", f"{age:.2f} v")
      r2.metric("Ennustettu pituus (PAH)", f"{pah:.1f} cm")
      r3.metric(
          "Pelioikeudellinen (Yli-ikäisyyssääntö)",
          eligibility_kr_str,
          delta=(
              "Täyttää vaatimuksen" if eligible_kr else "Ei täytä vaatimusta"
          ),
          delta_color="normal" if eligible_kr else "inverse",
      )

      row_to_save = [
          dom_fi,
          "Khamis-Roche",
          pallo_id,
          gender,
          round(age, 2),
          round(pah, 1),
          eligibility_kr_str,
      ]

    sheets_success, error_message = save_to_google_sheets(row_to_save)
    if sheets_success:
      st.success("Tiedot tallennettu onnistuneesti Google Sheets -taulukkoon!")
    else:
      st.error(
          f"Tallennus Google Sheetsiin epäonnistui. Virhe: {error_message}"
      )
