import streamlit as st
import vda_format


# Page configuration
st.set_page_config(
    page_title="PPAP Automotive Quality Hub",
    page_icon="🛫",
    layout="wide"
)


# Sidebar navigation
st.sidebar.header("Nawigacja")
selected_client = st.sidebar.selectbox(
    "Wybierz klienta",
    ["Strona główna", "VDA", "CNH", "MAN", "DAF"]
)


if selected_client == "Strona główna":
    # Home page
    st.subheader("Witaj w PPAP Automotive Quality Hub!")

    # Application status summary
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
elif selected_client == "VDA":
    # VDA documentation form
    vda_format.render_vda_ui()

else:
    # Information about planned modules
    st.subheader(f"Klient {selected_client} będzie wkrótce dostępny.")
    st.warning("Pracujemy nad dodaniem obsługi tego klienta. Prosimy o cierpliwość.")
