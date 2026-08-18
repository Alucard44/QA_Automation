import streamlit as st
import pandas as pd
from logika.obsluga_excel import generate_single_psw
from logika.generator_pdf import export_to_pdf


def _normalize_table_value(value):
    """Normalize a table value and return an empty string for missing entries."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.lower() in {"none", "nan", "<na>"}:
        return ""

    return text


def render_vda_ui():
    st.header("VDA PPA Report")

    # Select the PSW generation mode
    mass_production = st.checkbox("Tryb masowej produkcji PSW")

    st.sidebar.header("Dokumenty")
    st.sidebar.info("Aktualnie dostępne jest generowanie dokumentu PSW.")
    psw_check = st.sidebar.checkbox(
        "PSW",
        value=True,
        disabled=True,
        key="psw"
    )

    # Form submission requires clicking the generation button
    with st.form("vda_form", enter_to_submit=False):
        # VDA form header
        col_org_top, col_cust_top = st.columns(2)
        with col_org_top:
            st.markdown("**Organization**")
            org_name = st.text_input("Nazwa firmy (Org)", value="Example Automotive Organization", disabled=True)
            org_address = st.text_input("Adres (Org)", value="Example Street 1", disabled=True)
            org_zip_city = st.text_input("Kod pocztowy i miasto (Org)", value="00-000 Example City", disabled=True)
            org_country = st.text_input("Kraj (Org)", value="Poland", disabled=True)
        with col_cust_top:
            st.markdown("**Customer (recipient)**")
            cust_name = st.text_input("Nazwa firmy (Cust)", value="Example Customer")
            cust_address = st.text_input("Adres (Cust)", value="Customer Street 1")
            cust_zip_city = st.text_input("Kod pocztowy i miasto (Cust)", value="00-000 Customer City")
            cust_country = st.text_input("Kraj (Cust)", value="Poland")

        st.markdown("---")

        # Report reason and PPA procedure triggers
        if psw_check:
            st.markdown("**Reason for report creation**")
            col_rsn1, col_rsn2, col_rsn3 = st.columns(3)
            with col_rsn1:
                reason_ppa = st.checkbox("Report on production process and product approval (PPA)", value=True)
            with col_rsn2:
                reason_other = st.checkbox("Report on other samples")
            with col_rsn3:
                reason_requal = st.checkbox("Requalification")

            st.markdown("**Trigger of PPA procedure**")
            col_trig1, col_trig2, col_trig3 = st.columns(3)
            with col_trig1:
                trigger_sample = st.checkbox("Sample presentation")
                trigger_new_part = st.checkbox("New part", value=True)
                trigger_change_prod = st.checkbox("Changes to product")
            with col_trig2:
                trigger_change_proc = st.checkbox("Changes to production process")
                trigger_change_supply = st.checkbox("Change to supply chain")
            with col_trig3:
                trigger_reuse = st.checkbox("Re-use > 12 months standstill")
                trigger_updated = st.checkbox("Updated PPA documentation")

        st.markdown("---")

        # Part data
        if not mass_production:
            # Single generation mode data
            col_org, col_sam, col_cus = st.columns(3)
            with col_org:
                st.markdown("**Information about the organization**")
                report_number = st.text_input("Report number", value="0101/26")
                report_version = st.text_input("Report version", value="00")
                delivery_location = st.text_input("Delivery location", value="-")
                production_location = st.text_input("Production location", value="-")
                part_number = st.text_input("Part Number", value="123456678")
                part_name = st.text_input("Name", value="Demo Component")
                drawing_number = st.text_input("Drawing Number", value="323232323")
                version_date = st.text_input("Version / Date", value="01/01.01.2026")
            with col_sam:
                st.markdown("**Information about samples**")
                delivery_note_number = st.text_input("Delivery note number", value="121212312")
                delivery_quantity = st.text_input("Delivery quantity", value="5")
                batch_number = st.text_input("Batch number", value="1")
                sample_weight = st.text_input("Sample weight [kg]", value="0,5")
                hardware_version = st.text_input("Hardware version", value="-")
                diagnosis_status = st.text_input("Diagnosis Status", value="-")
                software_version = st.text_input("Software version", value="-")
                identification_duns = st.text_input("Identification / DUNS", value="111111")
            with col_cus:
                st.markdown("**Information about the customer**")
                customer = st.text_input("Customer", value="DEMO")
                order_number = st.text_input("Order Number PPA Sample", value="-")
                unloading_point = st.text_input("Unloading point", value="-")
                customer_part_number = st.text_input("Customer Part Number", value="CUST-PN-001")
                customer_part_name = st.text_input("Customer Name", value="Demo Component")
                customer_drawing_number = st.text_input("Customer drawing number", value="CUST-DWG-001")
                customer_version_date = st.text_input("Customer Version / Date", value="a / 01.12.2025")

            st.markdown("---")
            col_imds_check, col_imds_text = st.columns([1, 2])
            with col_imds_check:
                imds_check = st.checkbox("The IMDS record was created under the MDS ID No.:", value=True)
            with col_imds_text:
                imds_number = st.text_input("IMDS Number", value="11111111 / 1", label_visibility="collapsed")

        else:
            # Shared data for mass generation mode
            col_org, col_sam, col_cus = st.columns(3)
            with col_org:
                st.markdown("**Information about the organization**")
                delivery_location = st.text_input("Delivery location", value="-")
                production_location = st.text_input("Production location", value="-")
            with col_sam:
                st.markdown("**Information about samples**")
                delivery_note_number = st.text_input("Delivery note number", value="-")
                delivery_quantity = st.text_input("Delivery quantity", value="5")
                batch_number = st.text_input("Batch number", value="-")
                hardware_version = st.text_input("Hardware version", value="-")
                diagnosis_status = st.text_input("Diagnosis Status", value="-")
                software_version = st.text_input("Software version", value="-")
                identification_duns = st.text_input("Identification / DUNS", value="111111")
            with col_cus:
                st.markdown("**Information about the customer**")
                customer = st.text_input("Customer", value="DEMO")
                order_number = st.text_input("Order Number PPA sample", value="-")
                unloading_point = st.text_input("Unloading point", value="-")

            st.markdown("---")
            imds_check = st.checkbox("The IMDS record was created under the MDS ID No.: (Numery podaj w tabeli poniżej)", value=True)

        # Contact person details
        st.markdown("---")
        st.markdown("**Confirmation of organization (Osoba kontaktowa)**")
        col_cont1, col_cont2, col_cont3, col_cont4, col_cont5 = st.columns(5)
        with col_cont1:
            contact_name = st.text_input("Contact name", value="Demo User")
        with col_cont2:
            contact_department = st.text_input("Contact department", value="Quality")
        with col_cont3:
            contact_phone = st.text_input("Contact phone", value="+48 123 456 789")
        with col_cont4:
            contact_email = st.text_input("Contact email", value="demo.user@example.com")
        with col_cont5:
            contact_date = st.text_input("Date", value="01.01.2026")

        # Input table for mass generation mode
        if mass_production:
            st.markdown("---")
            st.subheader("Tryb masowy - Tabela Danych Zmiennych")
            variable_columns = [
                "Report number", "Report version", "Part Number", "Name",
                "Drawing number", "Version / Date", "Sample weight [kg]",
                "IMDS MDS ID No.", "Customer Part Number",
                "Customer Name", "Customer Drawing number", "Customer Version / Date"
            ]
            table_template = pd.DataFrame(columns=variable_columns)
            input_data = st.data_editor(table_template, num_rows="dynamic")

        submit_btn = st.form_submit_button("Generuj Dokumentację")

    # Prepare and pass data to the document generator
    if submit_btn:
        # Data shared by both generation modes
        organization_full_address = f"{org_name}\n{org_address}\n{org_zip_city}\n{org_country}"
        customer_full_address = f"{cust_name}\n{cust_address}\n{cust_zip_city}\n{cust_country}"

        form_data = {
            "Organization": organization_full_address,
            "Customer": customer_full_address,
            "Reason_PPA": reason_ppa,
            "Reason_Other": reason_other,
            "Reason_Requal": reason_requal,
            "Trigger_Sample": trigger_sample,
            "Trigger_NewPart": trigger_new_part,
            "Trigger_ChangeProd": trigger_change_prod,
            "Trigger_ChangeProc": trigger_change_proc,
            "Trigger_ChangeSupply": trigger_change_supply,
            "Trigger_Reuse": trigger_reuse,
            "Trigger_Updated": trigger_updated,
            "IMDS_Check": imds_check,
            "ContactName": contact_name,
            "ContactDepartment": contact_department,
            "ContactPhone": contact_phone,
            "ContactEmail": contact_email,
            "ContactDate": contact_date
        }

        if not mass_production:
            # Generate a single document
            form_data.update({
                "ReportNumber": report_number,
                "ReportVersion": report_version,
                "DeliveryLocation": delivery_location,
                "ProductionLocation": production_location,
                "PartNumber": part_number,
                "PartName": part_name,
                "DrawingNumber": drawing_number,
                "VersionDate": version_date,
                "DeliveryNoteNumber": delivery_note_number,
                "DeliveryQuantity": delivery_quantity,
                "BatchNumber": batch_number,
                "SampleWeight": sample_weight,
                "HardwareVersion": hardware_version,
                "DiagnosisStatus": diagnosis_status,
                "SoftwareVersion": software_version,
                "IdentificationDUNS": identification_duns,
                "Customer_Name": customer,
                "Customer_OrderNumber": order_number,
                "Customer_UnloadingPoint": unloading_point,
                "Customer_PartNumber": customer_part_number,
                "Customer_PartName": customer_part_name,
                "Customer_DrawingNumber": customer_drawing_number,
                "Customer_VersionDate": customer_version_date,
                "IMDS_Number": imds_number
            })


            with st.spinner("Generowanie raportu VDA w toku..."):
                try:
                    if psw_check:
                        psw_excel_path = generate_single_psw(form_data)
                        psw_pdf_path = export_to_pdf(psw_excel_path)
                        st.success(f"Sukces! Dokument wygenerowany jako: {psw_excel_path.name}")
                        st.success(f"Sukces! Dokument PDF wygenerowany jako: {psw_pdf_path.name}")

                except Exception as e:
                    st.error(f"Wystąpił błąd podczas generowania: {e}")
        else:
            # Generate documents in mass generation mode
            # Check whether the input table contains data
            if input_data.empty:
                st.warning("Tabela danych zmiennych jest pusta. Dodaj wiersze z numerami części!")
            else:
                generated_count = 0
                with st.spinner("Generowanie raportów masowych w toku..."):
                    # Process each row of the input table
                    for _, row in input_data.iterrows():
                        part_number = _normalize_table_value(row.get("Part Number"))

                        # Skip empty rows
                        if not part_number:
                            continue

                        row_form_data = form_data.copy()

                        # Prepare data for the current row
                        row_form_data.update({
                            # Data shared by all rows
                            "DeliveryLocation": delivery_location,
                            "ProductionLocation": production_location,
                            "Customer_Name": customer,
                            "Customer_OrderNumber": order_number,
                            "Customer_UnloadingPoint": unloading_point,
                            "DeliveryNoteNumber": delivery_note_number,
                            "DeliveryQuantity": delivery_quantity,
                            "BatchNumber": batch_number,
                            "HardwareVersion": hardware_version,
                            "DiagnosisStatus": diagnosis_status,
                            "SoftwareVersion": software_version,
                            "IdentificationDUNS": identification_duns,

                            # Data read from the current table row
                            "ReportNumber": _normalize_table_value(row.get("Report number")),
                            "ReportVersion": _normalize_table_value(row.get("Report version")),
                            "PartNumber": part_number,
                            "PartName": _normalize_table_value(row.get("Name")),
                            "DrawingNumber": _normalize_table_value(row.get("Drawing number")),
                            "VersionDate": _normalize_table_value(row.get("Version / Date")),
                            "SampleWeight": _normalize_table_value(row.get("Sample weight [kg]")),
                            "IMDS_Number": _normalize_table_value(row.get("IMDS MDS ID No.")),
                            "Customer_PartNumber": _normalize_table_value(row.get("Customer Part Number")),
                            "Customer_PartName": _normalize_table_value(row.get("Customer Name")),
                            "Customer_DrawingNumber": _normalize_table_value(row.get("Customer Drawing number")),
                            "Customer_VersionDate": _normalize_table_value(row.get("Customer Version / Date"))
                        })

                        try:
                            if psw_check:
                                psw_excel_path = generate_single_psw(row_form_data)
                                export_to_pdf(psw_excel_path)

                            generated_count += 1
                        except Exception as e:
                            st.error(f"Błąd przy generowaniu dla PN {part_number}: {e}")

                st.success(f"Sukces! Wygenerowano masowo {generated_count} plików PSW na podstawie danych z tabeli!")
