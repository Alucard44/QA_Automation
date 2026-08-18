import streamlit as st
import vda_format


# Konfiguracja strony
st.set_page_config(
    page_title="PPAP Automotive Quality Hub",
    page_icon="🛫",
    layout="wide"
)


# Nawigacja w panelu bocznym
st.sidebar.header("Nawigacja")
klient = st.sidebar.selectbox(
    "Wybierz klienta",
    ["Strona główna", "VDA", "CNH", "MAN", "DAF"]
)


if klient == "Strona główna":
    # Strona główna
    st.subheader("Witaj w PPAP Automotive Quality Hub!")

    # Podsumowanie stanu aplikacji
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Obsługiwani klienci", value="1 (VDA)")
    with col2:
        st.metric(label="Dostępne dokumenty", value="PSW")
    with col3:
        st.metric(label="Wersja aplikacji", value="1.0.0")

    st.divider()

    st.info(
        "👈 Wybierz klienta z menu po lewej stronie. "
        "Dla klientów bez dedykowanej konfiguracji wybierz VDA."
    )

    st.markdown("""
    ## O aplikacji
    1. **Wybierz klienta** z menu po lewej stronie. Dla klientów bez dedykowanej konfiguracji wybierz VDA.
    2. **Wybierz tryb generowania** pojedynczy lub masowy.
    3. **Uzupełnij formularz** wymaganymi danymi.
    4. **Kliknij „Generuj Dokumentację”**, aby utworzyć pliki Excel i PDF.
    """)
elif klient == "VDA":
    # Formularz dokumentacji VDA
    vda_format.render_vda_ui()

else:
    # Informacja o planowanych modułach
    st.subheader(f"Klient {klient} będzie wkrótce dostępny.")
    st.warning("Pracujemy nad dodaniem obsługi tego klienta. Prosimy o cierpliwość.")