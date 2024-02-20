"""
 * Copyright (C) 2023 SFTI and Swedish Local Authorities and Regions (SALAR)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
  limitations under the License.
 """
import openpyxl
import re
import xml.etree.ElementTree as el_sbdh_tree
import uuid
from datetime import datetime


def load_code_list(wb, col_range):
    ws = wb["CodeLists"]
    code_list = {}

    start_col, end_col = col_range.split(":")

    # Convert column letters to indices
    min_index = sum((ord(char.upper()) - ord("A") + 1) * (26 ** i) for i, char in enumerate(reversed(start_col)))
    max_index = sum((ord(char.upper()) - ord("A") + 1) * (26 ** i) for i, char in enumerate(reversed(end_col)))

    number_of_columns = max_index - min_index + 1
    for row in ws.iter_rows(min_row=3, min_col=min_index, max_col=max_index):
        attr1_value = ""
        attr2_value = ""
        attr3_value = ""

        if number_of_columns == 2:
            cell_name, cell_code = row
        elif number_of_columns == 3:
            cell_name, cell_code, cell_attr1 = row
            attr1_value = cell_attr1.value
        elif number_of_columns == 4:
            cell_name, cell_code, cell_attr1, cell_attr2 = row
            attr1_value = cell_attr1.value
            attr2_value = cell_attr2.value
        elif number_of_columns == 5:
            cell_name, cell_code, cell_attr1, cell_attr2, cell_attr3 = row
            attr1_value = cell_attr1.value
            attr2_value = cell_attr2.value
            attr3_value = cell_attr3.value
        else:
            raise ValueError(f"Codelist with incorrect range: {col_range} = {number_of_columns}")

        # Skip rows where either the code or the name is empty
        if cell_code.value is None or cell_name.value is None:
            break

        # Add to dictionary
        code_list[str(cell_name.value)] = {"Code": cell_code.value, "Attr1": attr1_value, "Attr2": attr2_value, "Attr3": attr3_value}

    return code_list


def header_cell(config, name):
    try:
        return str(config.get("HeaderCell", name))
    except Exception as e:
        raise Exception(f"Header cell for {name} not found in the configuration.")


def cl_range(config, name):
    try:
        return str(config.get("CodeLists", name))
    except Exception as e:
        raise Exception(f"Code list range for {name} not found in the configuration.")


def col_index(config, name):
    try:
        return int(config.get("LineColIndex", name))
    except Exception as e:
        raise Exception(f"Column index range for {name} not found in the configuration.")


def add_element(el_tree, parent_element, element_name: str, element_value: str):
    if not is_cell_empty(element_value):
        c = el_tree.SubElement(parent_element, element_name)
        c.text = element_value
        return c


def add_attribute(parent_element, attribute_name: str, attribute_value: str):
    if not is_cell_empty(attribute_value) and parent_element is not None:
        parent_element.set(attribute_name, attribute_value)


def add_item_classification(el_tree, parent_element, name, list_id, list_version_id, value: str):
    if not is_cell_empty(value):
        cac = el_tree.SubElement(parent_element, "cac:CommodityClassification")
        c = add_element(cac, "cbc:ItemClassificationCode", value)
        add_attribute(c, "listID", list_id)
        add_attribute(c, "name", name)
        add_attribute(c, "listVersionID", list_version_id)


def add_additional_item_prop(el_tree, parent_element, name, name_code, name_code_list_id, value, value_qualifier):
    if not is_cell_empty(name):
        cac = el_tree.SubElement(parent_element, "cac:AdditionalItemProperty")
        add_element(el_tree, cac, "cbc:Name", name)
        c = add_element(el_tree, cac, "cbc:NameCode", name_code)
        add_attribute(c, "listID", name_code_list_id)
        add_element(el_tree, cac, "cbc:Value", value)
        add_element(el_tree, cac, "cbc:ValueQualifier", value_qualifier)


def add_item_certificate(el_tree, parent_element, label_name_id, certificate_type, remark, qualifier_code):
    if not is_cell_empty(label_name_id):
        cac = el_tree.SubElement(parent_element, "cac:Certificate")
        add_element(el_tree, cac, "cbc:ID", label_name_id)
        add_element(el_tree, cac, "cbc:CertificateTypeCode", "NA")
        add_element(el_tree, cac, "cbc:CertificateType", certificate_type)
        add_element(el_tree, cac, "cbc:Remarks", remark)
        cac1 = el_tree.SubElement(cac, "cac:IssuerParty")
        cac2 = el_tree.SubElement(cac1, "cac:PartyName")
        add_element(el_tree, cac2, "cbc:Name", "NA")
        if qualifier_code != "":
            cac3 = el_tree.SubElement(cac, "cac:DocumentReference")
            add_element(el_tree, cac3, "cbc:ID", qualifier_code)


def add_item_dimension(el_tree, parent_element, attribute_id, measure, minimum_measure, maximum_measure, unit_code):
    cac = el_tree.SubElement(parent_element, "cac:Dimension")
    add_element(el_tree, cac, "cbc:AttributeID", attribute_id)
    c = add_element(el_tree, cac, "cbc:Measure", measure)
    add_attribute(c, "unitCode", unit_code)
    c1 = add_element(el_tree, cac, "cbc:MinimumMeasure", minimum_measure)
    add_attribute(c1, "unitCode", unit_code)
    c2 = add_element(el_tree, cac, "cbc:MaximumMeasure", maximum_measure)
    add_attribute(c2, "unitCode", unit_code)


def add_price(el_tree, parent_element, minimum_quantity, minimum_quantity_unit_code, price_amount, currency_id, base_quantity, base_quantity_unit_code, price_type, start_date, end_date, lead_time):
    cac = el_tree.SubElement(parent_element, "cac:RequiredItemLocationQuantity")
    c = add_element(el_tree, cac, "cbc:LeadTimeMeasure", lead_time)
    add_attribute(c, "unitCode", "DAY")
    c = add_element(el_tree, cac, "cbc:MinimumQuantity", minimum_quantity)
    add_attribute(c, "unitCode", minimum_quantity_unit_code)
    price = el_tree.SubElement(cac, "cac:Price")
    c = add_element(el_tree, price, "cbc:PriceAmount", price_amount)
    add_attribute(c, "currencyID", currency_id)
    c = add_element(el_tree, price, "cbc:BaseQuantity", base_quantity)
    add_attribute(c, "unitCode", base_quantity_unit_code)
    add_element(el_tree, price, "cbc:PriceType", price_type)
    if not is_cell_empty(start_date) or not is_cell_empty(end_date):
        cac = el_tree.SubElement(price, "cac:ValidityPeriod")
        add_element(el_tree, cac, "cbc:StartDate", start_date.split(" ")[0])
        add_element(el_tree, cac, "cbc:EndDate", end_date.split(" ")[0])


def get_code(key, codelist, alternate_field=None) -> str:
    value = ""
    if alternate_field is None:
        value = codelist.get(key, {}).get("Code", "")
    else:
        value = codelist.get(key, {}).get(alternate_field, "")
    return value


def is_cell_empty(value: str) -> bool:
    return normalize_space(value) == "" or value == "None"


def normalize_space(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()


def separated_string(input_str):
    # split string into array
    values = input_str.split(";")

    # handle trailing ;
    if values[-1] == "":
        values.pop()

    return values


def check_spreadsheet_consistency(wb: openpyxl.Workbook):
    try:
        #Try to assign the sheets to ensure they exist
        sheet_header = wb["CatalogueHeader"]
        sheet_lines = wb["CatalogueLines"]
        sheet_codelists = wb["CodeLists"]
    except Exception as e:
        raise ValueError("Excel sheet names not according to SFTI template.") from e

    # Check that the columns have not been rearranged
    for col_idx, col in enumerate(sheet_lines.iter_cols(min_row=1, max_row=1), 1):
        for cell in col:
            if not cell.value:
                return

            # Check if the cell value matches the column index
            if str(cell.value) != str(col_idx):
                raise ValueError(f"Column index not correct for column {str(col_idx)}. Ensure that the SFTI template has not been altered with.")


def add_to_sbdh(catalogue, sender_id_scheme: str, sender_id: str, receiver_id_scheme: str, receiver_id: str, sender_countrycode: str):

    sbdh = el_sbdh_tree.Element("StandardBusinessDocument")
    sbdh.set("xmlns", "http://www.unece.org/cefact/namespaces/StandardBusinessDocumentHeader")

    h = el_sbdh_tree.SubElement(sbdh, "StandardBusinessDocumentHeader")
    add_element(el_sbdh_tree,h,"HeaderVersion","1.0")
    e = el_sbdh_tree.SubElement(h, "Sender")
    e = add_element(el_sbdh_tree,e ,"Identifier",f"{sender_id_scheme}:{sender_id}")
    add_attribute(e, "Authority", "iso6523-actorid-upis")
    e = el_sbdh_tree.SubElement(h, "Receiver")
    e = add_element(el_sbdh_tree, e,"Identifier",f"{receiver_id_scheme}:{receiver_id}")
    add_attribute(e, "Authority", "iso6523-actorid-upis")
    d = el_sbdh_tree.SubElement(h, "DocumentIdentification")
    e = add_element(el_sbdh_tree, d, "Standard", "urn:oasis:names:specification:ubl:schema:xsd:Catalogue-2")
    e = add_element(el_sbdh_tree, d, "TypeVersion", "2.1")
    e = add_element(el_sbdh_tree, d, "InstanceIdentifier", str(uuid.uuid4()))
    e = add_element(el_sbdh_tree, d, "Type", "Catalogue")
    e = add_element(el_sbdh_tree, d, "CreationDateAndTime", datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))
    b = el_sbdh_tree.SubElement(h, "BusinessScope")
    bs = el_sbdh_tree.SubElement(b, "Scope")
    e = add_element(el_sbdh_tree, bs, "Type", "DOCUMENTID")
    e = add_element(el_sbdh_tree, bs, "InstanceIdentifier", "urn:oasis:names:specification:ubl:schema:xsd:Catalogue-2::Catalogue##urn:fdc:peppol.eu:poacc:trns:catalogue:3::2.1")
    e = add_element(el_sbdh_tree, bs, "Identifier","busdox-docid-qns")

    bs = el_sbdh_tree.SubElement(b, "Scope")
    e = add_element(el_sbdh_tree, bs, "Type", "PROCESSID")
    e = add_element(el_sbdh_tree, bs, "InstanceIdentifier", "urn:fdc:peppol.eu:poacc:bis:catalogue_wo_response:3")
    e = add_element(el_sbdh_tree, bs, "Identifier","cenbii-procid-ubl")

    bs = el_sbdh_tree.SubElement(b, "Scope")
    e = add_element(el_sbdh_tree, bs, "Type", "COUNTRY_C1")
    e = add_element(el_sbdh_tree, bs, "InstanceIdentifier", sender_countrycode)

    sbdh.append(catalogue)

    return sbdh



