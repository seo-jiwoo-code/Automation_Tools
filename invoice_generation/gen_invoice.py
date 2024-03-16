from docx_lib import *
import pandas as pd
import numpy as np
import json

from config import *

def format_cost(cost_int):
    if cost_int < 0:
        return "-${:.2f}".format(abs(float(cost_int)))
    return "${:.2f}".format(float(cost_int))

def add_table_total_cost(table, total_cost):
    table.cell(1, len(table.columns)-1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    table.cell(1, len(table.columns)-1).merge(table.cell(len(table.rows)-1, len(table.columns)-1))
    # print(table.cell(1, len(table.columns)-1).text)
    table.cell(1, len(table.columns)-1).text = format_cost(total_cost)
    set_table_cellfont(table, 8)
    set_table_cellalignment(table)

def add_logo(doc):  
    logo_path = "logo.png"
    pic = add_picture(doc, logo_path, width=Cm(4.13), height=Cm(1.23))
    align_para(pic, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    # doc.add_section(WD_SECTION.ODD_PAGE)
    # print(pic._p.xml)

def gen_heading1(doc, text):
    heading1 = add_heading(doc, 1, text)
    set_text_font(heading1, font_name = "Times New Roman", font_size = 15)
    set_text_alignment(heading1, left_indent=0, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    set_text_underline(heading1)
    set_para_spacing(heading1, before=0, after=10)

def gen_heading2(doc, text):
    heading2 = add_heading(doc, 3, text)
    set_text_font(heading2, font_name = "Times New Roman", font_size = 12)
    set_text_alignment(heading2, left_indent=0, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_para_spacing(heading2, before=0, after=5)

def gen_table(doc, data, header_type="row"):
    table = add_table(doc, data)
    num_rows = len(data)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_colwidth(table)
    set_table_width(table)
    set_table_borders(table)
    set_table_rowheight(table, num_rows)
    set_table_cellpadding(table, padding=100)
    set_table_headerbgcolor(table, bgcolor = 'C7C7C7', header_type=header_type)
    set_table_header_italic_bold(table, header_type=header_type)
    set_table_cellfont(table, 8)
    set_table_cellalignment(table)
    set_table_alignment(table)
    prevent_table_rowbreak(table, num_rows)
    return table

### ADD JSON TO DOCX

def add_std_test_details(doc, std_test_details_json):
    gen_heading1(doc, "Standard Test Details")
    col_map = {
        TEST_PARAM_COL  : "Test Parameters",
        METHOD_REFERENCE_COL : "Method Reference",
        "TAT (C.Days)" : "TAT (calendar days)",
        "Min Sample (g)" : "Min. Sample Size (g)",
        UNITS_OF_MEASUREMENT_COL : "Units of Measurement"
    }
    for test_type in std_test_details_json:
        gen_heading2(doc, f"{test_type}".replace("\n", ""))
        table_header = ["No."] + [col_map[key] for key in col_map]
        table_data  = [table_header]
        for itr, test_param_detail in enumerate(std_test_details_json[test_type]):
            table_row = [itr+1]
            for key in col_map:
                if type(test_param_detail[key]) == str and test_param_detail[key].isnumeric():
                    table_row.append(int(test_param_detail[key]))
                elif type(test_param_detail[key]) == float:
                    table_row.append(int(test_param_detail[key]))
                else:
                    table_row.append(test_param_detail[key])
            table_data.append(table_row)
        gen_table(doc, table_data)
        doc.add_paragraph("")

def add_shelf_life_test_details(doc, shelf_life_test_details_json):
    gen_heading1(doc, "Shelf-Life Test Details")
    bottom_table = ["Samples Mass Required", "Acceleration", "Turnaround Time"]
    for shelf_test in shelf_life_test_details_json:
        gen_heading2(doc, f"{shelf_test['Shelf-Life Header']}".replace("\n", ""))
        table_data  = []
        for key in shelf_test['Shelf-Life Detail']:
            if key not in bottom_table:
                table_row = [key, shelf_test['Shelf-Life Detail'][key]]
                table_data.append(table_row)

        empty_row = ["", ""]
        table_data.append(empty_row)
        merge_row_id = len(table_data)-1

        for key in shelf_test['Shelf-Life Detail']:
            if key in bottom_table:
                table_row = [key, shelf_test['Shelf-Life Detail'][key]]
                table_data.append(table_row)

        table = gen_table(doc, table_data, header_type="col")
        table.cell(merge_row_id, 0).merge(table.cell(merge_row_id, 1))
        # table.cell(merge_row_id, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_color(table.cell(merge_row_id, 0), color="FFFFFF")

        doc.add_paragraph("")

def add_std_test_charges(doc, std_test_charges_json):
    col_map = {
        LAB_TEST_CATEGORY_COL : "Tests Category",
        TEST_PARAM_COL : "Test Parameters",
        NUMBER_OF_SAMPLES_COL : "No. of Sample(s)",
        SAMPLE_LABELS_COL : "Sample Labels",
        SAMPLE_DESCRIPTION_COL : "Sample Description & Mass",
    }
    gen_heading1(doc, "Standard Test Charges (inc. tax)")
    table_header = ["No."] + [col_map[key] for key in col_map] + ["Cost (SGD)"]
    table_data  = [table_header]
    for itr, std_test_charge in enumerate(std_test_charges_json):
        table_row = [itr+1]
        for key in col_map:
            if type(std_test_charge[key]) == str and std_test_charge[key].isnumeric():
                table_row.append(int(std_test_charge[key]))
            elif type(std_test_charge[key]) == float:
                table_row.append(int(std_test_charge[key]))
            else:
                table_row.append(std_test_charge[key])
        table_row.append(format_cost(std_test_charge[COST_COL]))
        table_data.append(table_row)
    gen_table(doc, table_data)
    doc.add_paragraph("")

def add_shelf_life_test_charges(doc, shelf_life_test_charges_json):
    col_map = {
        "Shelf-Life Description" : "Shelf-Life Description",
        "Shelf-Life Ref. No." : "Shelf-Life Ref. No.",
        "Total No. of Sub-sample(s)" : "Total No. of Sub-sample(s)",
        "Sub-sample Label(s)" : "Sub-sample Label(s)",
        "SKU Temp., Description & Mass" : "SKU Temp., Description & Mass",
    }
    gen_heading1(doc, "Shelf-Life Tests Charges (inc. tax)") 
    table_header = ["No."] + [col_map[key] for key in col_map] + ["Cost (SGD)"]
    table_data  = [table_header]
    for itr, shelf_life_test_charge in enumerate(shelf_life_test_charges_json):
        table_row = [itr+1]
        for key in col_map:
            table_row.append(shelf_life_test_charge[key])
        table_row.append(format_cost(shelf_life_test_charge[COST_COL]))
        table_data.append(table_row)
    gen_table(doc, table_data)
    doc.add_paragraph("")

def add_total_lab_test_charges(doc, total_lab_test_charges_json):
    gen_heading1(doc, "Total Lab Test Charges (inc. tax)")
    col_map = {
        "Description" : "Description",
    }
    table_header = ["No."] + [key for key in col_map] + ["Cost (SGD)", "Total Cost (SGD)"]
    total_cost = 0
    table_data  = [table_header]
    for idx, data in enumerate(total_lab_test_charges_json):
        table_row = [idx+1] + [data[key] for key in col_map] + [format_cost(data[COST_COL]), ""]
        total_cost += data[COST_COL]
        table_data.append(table_row)

    table = gen_table(doc, table_data)
    add_table_total_cost(table, total_cost)
    doc.add_paragraph("")

def add_total_food_service_charges(doc, food_services_test_charges_json):
    gen_heading1(doc, "Total Food Service Charges (inc. tax)")
    col_map = {
        "Description" : "Description",
        "No. of Food Products" : "No. of Food Products",
        "Product Description" : "Product Description",
    }
    table_header = ["No."] + [key for key in col_map] + ["Cost (SGD)", "Total Cost (SGD)"]
    total_cost = 0
    table_data  = [table_header]
    for idx, data in enumerate(food_services_test_charges_json):
        table_row = [idx+1] + [data[key] for key in col_map] + [format_cost(data[COST_COL]), ""]
        total_cost += data[COST_COL]
        table_data.append(table_row)

    table = gen_table(doc, table_data)
    add_table_total_cost(table, total_cost)
    doc.add_paragraph("")

def add_total_cost(doc, total_cost_json):
    gen_heading1(doc, "Total Cost (inc. tax)")
    col_map = {
        "Description" : "Description",
    }
    table_header = ["No."] + [key for key in col_map] + ["Cost (SGD)", "Total Cost (SGD)"]
    total_cost = 0
    table_data  = [table_header]

    for idx, data in enumerate(total_cost_json):
        cost_str = format_cost(data[COST_COL])
        table_row = [idx+1] + [data[key] for key in col_map] + [cost_str, ""]
        total_cost += data[COST_COL]
        table_data.append(table_row)

    table = gen_table(doc, table_data)
    add_table_total_cost(table, total_cost)
    doc.add_paragraph("")

def replace_with_todo(text):
    if text == "":
        return "<TODO>"
    return str(text)

def create_std_test_details_json(input_json, lab_df):
    input_json = input_json["Samples"]
    # If Empty Test Details
    if len(input_json) == 0:
        std_test_detail = { col : "<TODO>" for col in [TEST_PARAM_COL, "Method Reference", "TAT (C.Days)", "Min Sample (g)", UNITS_OF_MEASUREMENT_COL] }   
        std_test_details_data = { "<TODO>" : [std_test_detail]}
        std_test_details_data = {}
        return std_test_details_data
    
    std_test_details_data = {}
    for test_param_detail in input_json:
        test_param_df = lab_df.copy()
        for col in [
            (TEST_PARAM_COL, TEST_PARAM_COL),
            (LAB_ID_COL, LAB_ID_COL),
            (LAB_TEST_CATEGORY_COL, LAB_TEST_CATEGORY_COL),
            (METHOD_REFERENCE_CODE_COL, MRC_COL)
            ]:
            if test_param_detail[col[0]]:
                test_param_df = test_param_df[test_param_df[col[1]] == test_param_detail[col[0]]]

        if len(test_param_df) == 0:
            std_test_detail = { col : "<TODO>" for col in [TEST_PARAM_COL, "Method Reference", "TAT (C.Days)", "Min Sample (g)", UNITS_OF_MEASUREMENT_COL] }
            std_test_detail[TEST_PARAM_COL] = test_param_detail[TEST_PARAM_COL]
            test_type = replace_with_todo(test_param_detail[LAB_TEST_CATEGORY_COL])
            print(f"No params found for {json.dumps(test_param_detail, indent=4 )}")
        elif len(test_param_df) > 1:
            std_test_detail = { col : "<TODO>" for col in [TEST_PARAM_COL, "Method Reference", "TAT (C.Days)", "Min Sample (g)", UNITS_OF_MEASUREMENT_COL]}
            std_test_detail[TEST_PARAM_COL] = test_param_detail[TEST_PARAM_COL]
            test_type = replace_with_todo(test_param_detail[LAB_TEST_CATEGORY_COL])
            print(f"Multiple test params found for {json.dumps(test_param_detail, indent=4)}")
            print(test_param_df)
        else:
            std_test_detail = { col : test_param_df[col].iloc[0] for col in [TEST_PARAM_COL, "Method Reference", "TAT (C.Days)", "Min Sample (g)", UNITS_OF_MEASUREMENT_COL] }
            test_type = replace_with_todo(test_param_df[LAB_TEST_CATEGORY_COL].iloc[0])
            std_test_detail[TEST_PARAM_COL] = test_param_detail[TEST_PARAM_COL]
            # print("Test Detail", TEST_PARAM_COL)
            # print(std_test_detail)

        if test_type in std_test_details_data:
            if std_test_detail not in std_test_details_data[test_type]:
                std_test_details_data[test_type] += [std_test_detail]
            else:
                print("Duplicate Test Detail", std_test_detail)
        else:
            std_test_details_data[test_type] = [std_test_detail]

    return std_test_details_data

def create_std_test_charges_json(input_json, lab_df):
    input_json = input_json["Samples"]
    # If Empty Test Details
    if len(input_json) == 0:
        std_test_charge = {
            LAB_TEST_CATEGORY_COL : ["<TODO>"],
            TEST_PARAM_COL : ["<TODO>"],
            SAMPLE_LABELS_COL : ["<TODO>"],
            NUMBER_OF_SAMPLES_COL : "<TODO>",
            SAMPLE_DESCRIPTION_COL : "<TODO>",
            COST_COL : np.nan
        } 
        std_test_charges_data = [std_test_charge]
        return std_test_charges_data
    
    std_test_charges_data = []
    for test_param_detail in input_json:
        lab_test_category = test_param_detail[LAB_TEST_CATEGORY_COL]
        test_param = test_param_detail[TEST_PARAM_COL]
        sample_label = test_param_detail[SAMPLE_LABELS_COL]

        flag = True
        for itr in range(len(std_test_charges_data)):
            if lab_test_category in std_test_charges_data[itr][LAB_TEST_CATEGORY_COL] and test_param in std_test_charges_data[itr][TEST_PARAM_COL]:
                if sample_label not in std_test_charges_data[itr][SAMPLE_LABELS_COL]:
                    std_test_charges_data[itr][SAMPLE_LABELS_COL] += [sample_label]
                std_test_charges_data[itr][NUMBER_OF_SAMPLES_COL] += test_param_detail[NUMBER_OF_SAMPLES_COL]
                std_test_charges_data[itr][COST_COL] += test_param_detail[COST_COL]
                flag = False
            elif lab_test_category in std_test_charges_data[itr][LAB_TEST_CATEGORY_COL] and sample_label in std_test_charges_data[itr][SAMPLE_LABELS_COL]:
                std_test_charges_data[itr][TEST_PARAM_COL] += [test_param]
                std_test_charges_data[itr][NUMBER_OF_SAMPLES_COL] += test_param_detail[NUMBER_OF_SAMPLES_COL]
                std_test_charges_data[itr][COST_COL] += test_param_detail[COST_COL]
                flag = False

        if flag:
            std_test_charge = {
                LAB_TEST_CATEGORY_COL : [lab_test_category],
                TEST_PARAM_COL : [test_param],
                SAMPLE_LABELS_COL : [sample_label],
                NUMBER_OF_SAMPLES_COL : test_param_detail[NUMBER_OF_SAMPLES_COL],
                SAMPLE_DESCRIPTION_COL : test_param_detail[SAMPLE_DESCRIPTION_COL],
                COST_COL : test_param_detail[COST_COL]
            }
            std_test_charges_data.append(std_test_charge)

    return std_test_charges_data

def create_shelf_life_test_charges_json(input_json):
    shelf_life_list = input_json["Shelf-Life"]
    shelf_life_test_charges_json = []
    for shelf_life_dict in shelf_life_list:
        print(shelf_life_dict)
        tot_num_sub_samples = f"{shelf_life_dict['No. of Test Dates']} x {shelf_life_dict['No. of SKUs']} = {shelf_life_dict['No. of Test Dates'] * shelf_life_dict['No. of SKUs']}"
        shelf_life_test_charge = {
            "Shelf-Life Description" : shelf_life_dict["Header name"],
            "Shelf-Life Ref. No." : shelf_life_dict["Shelf-Life Ref. No."],
            "Total No. of Sub-sample(s)" : tot_num_sub_samples,
            "Sub-sample Label(s)" : shelf_life_dict["Sub-sample Label(s)"],
            "SKU Temp., Description & Mass" : shelf_life_dict["SKU Temp., Description & Mass"],
            "Cost" : float(shelf_life_dict[COST_COL])
        }
        shelf_life_test_charges_json.append(shelf_life_test_charge)

    return shelf_life_test_charges_json

def create_shelf_life_test_details_json(input_json):
    shelf_life_list = input_json["Shelf-Life"]
    allowed_cols = [
        "Shelf-Life Ref. No.",
        "Type of Shelf-Life",
        "No. of Test Dates",
        "Storage Temperature",
        "Testing Dates",
        "No. of Test Parameters",
        "Test Parameters",
        "No. of SKUs",
        "Detailed Test Parameters",
        "Samples Mass Required",
        "Acceleration",
        "Turnaround Time"
    ]
    shelf_life_test_details_json = []
    for shelf_life_dict in shelf_life_list:
        shelf_life_test_detail = {
            "Shelf-Life Header" : shelf_life_dict["Header name"],
            "Shelf-Life Detail" : {key: shelf_life_dict[key] for key in allowed_cols}
        }
        shelf_life_test_details_json.append(shelf_life_test_detail)

    return shelf_life_test_details_json

def create_total_lab_test_charges_json(shelf_life_test_charges_json, std_test_charges_json):
    total_lab_test_charges_json = []

    if std_test_charges_json:
        total_std_test_charges = 0
        for std_test_charge in std_test_charges_json:
            total_std_test_charges += float(std_test_charge[COST_COL])
    else:
        total_std_test_charges = 0

    total_lab_test_charges_json.append({
        "Description" : "Standard Test Charges",
        "Cost" : total_std_test_charges,
    })

    if shelf_life_test_charges_json:
        total_shelf_life_test_charges = 0
        for shelf_life_test_charge in shelf_life_test_charges_json:
            total_shelf_life_test_charges += float(shelf_life_test_charge[COST_COL])
    else:
        total_shelf_life_test_charges = 0
    
    total_lab_test_charges_json.append({
        "Description" : "Shelf-Life Test Charges",
        "Cost" : total_shelf_life_test_charges,
    })

    return total_lab_test_charges_json

def create_total_cost_json(invoice, total_lab_test_charges_json, food_services_test_charges_json):
    total_cost_json = []

    if total_lab_test_charges_json:
        total_total_lab_test_charges = 0
        for total_lab_test_charge in total_lab_test_charges_json:
            total_total_lab_test_charges += float(total_lab_test_charge[COST_COL])
        total_cost_json.append({
            "Description" : "Total Lab Test Charges",
            "Cost" : total_total_lab_test_charges,
        })
    
    if food_services_test_charges_json:
        total_food_services_test_charges = 0
        for food_services_test_charge in food_services_test_charges_json:
            total_food_services_test_charges += float(food_services_test_charge[COST_COL])
        total_cost_json.append({
            "Description" : "Total Food Service Charges",
            "Cost" : total_food_services_test_charges,
        })

    if invoice["Other Costs"]["Discount"]:
        total_cost_json.append({
            "Description" : "Discount",
            "Cost" : -invoice["Other Costs"]["Discount"],
        })

    # total_cost_json.append({
    #     "Description" : "Discount",
    #     "Cost" : np.nan,
    # })

    return total_cost_json

def create_food_service_test_charges_json(input_json):
    food_services_test_charges_json = [{
        "Description" : "NA",
        "No. of Food Products" : 0,  
        "Product Description" : "NA",
        "Cost" : 0,
    }]
    return food_services_test_charges_json

def parse_json_input(input_json, lab_df):
    if len(input_json) == 0:
        lab_test_charge_data = [[1] + ["<TODO>"]*6]
        lab_test_detail = [1] + ["<TODO>"]*5
        lab_test_details_data = {
            "<TODO>" : [lab_test_detail]
        }
        return lab_test_details_data, lab_test_charge_data
    
    lab_test_charge_data = []
    lab_test_details_data = {}
    for idx0, sample in enumerate(input_json):
        # If empty Sample Label, Sample Description, Test Param Details
        sample[SAMPLE_LABELS_COL] = [replace_with_todo(ele) for ele in sample[SAMPLE_LABELS_COL]]
        sample[SAMPLE_DESCRIPTION_COL] = replace_with_todo(sample[SAMPLE_DESCRIPTION_COL])
        if len(sample[TEST_PARAM_DETAILS_COL]) == 0:
            sample_charges = {
                TEST_PARAMS_COL : ["<TODO>", "<TODO>"],
                LAB_TEST_CATEGORY_COL : ["<TODO>", "<TODO>"]
            }
            lab_test_detail = [test_param, "<TODO>", "<TODO>", "<TODO>", "<TODO>"]
            test_type = "<TODO>"
            if test_type in lab_test_details_data:
                lab_test_detail = [len(lab_test_details_data[test_type])+1] + lab_test_detail
                lab_test_details_data[test_type] += [lab_test_detail]
            else:
                lab_test_details_data[test_type] = {}
                lab_test_detail = [1] + lab_test_detail
                lab_test_details_data[test_type] = [lab_test_detail]
            continue
        
        # Else
        sample_charges = {
            TEST_PARAMS_COL : set(),
            LAB_TEST_CATEGORY_COL : set(),
        }

        for idx1, test_param_data in enumerate(sample[TEST_PARAM_DETAILS_COL]):
            # If empty Test Param, Lab ID
            test_param = replace_with_todo(test_param_data[TEST_PARAM_COL])
            lab_id = replace_with_todo(test_param_data[LAB_ID_COL])
            test_param_df = lab_df[lab_df[TEST_PARAM_COL] == test_param]
            # for filter in ["Lab ID", ""]

            test_param_df = test_param_df[test_param_df[LAB_ID_COL] == lab_id]
            sample_charges[TEST_PARAMS_COL].add(test_param)

            # if len(test_param_df) == 0:
            lab_test_detail = [test_param, "<TODO>", "<TODO>", "<TODO>", "<TODO>"]
            test_type = "<TODO>"

            print("TEST PARAM DF", test_param_df.dtypes)

            for index, row in test_param_df.iterrows():
                # If empty Test ID, Test Type
                test_type = replace_with_todo(row['Test Type NotNull'])
                lab_test_detail = [
                    test_param,
                    replace_with_todo(row['Method Reference']),
                    replace_with_todo(row['TAT']),
                    replace_with_todo(row['Minimum Sample']),
                    replace_with_todo(row['Units of Measurement'])
                ]
                break

            sample_charges[LAB_TEST_CATEGORY_COL].add(test_type)
            
            if test_type in lab_test_details_data:
                lab_test_detail = [len(lab_test_details_data[test_type])+1] + lab_test_detail
                lab_test_details_data[test_type] += [lab_test_detail]
            else:
                lab_test_details_data[test_type] = {}
                lab_test_detail = [1] + lab_test_detail
                lab_test_details_data[test_type] = [lab_test_detail]

        num_samples = len(sample[SAMPLE_LABELS_COL])
        lab_test_charge_data.append(
            [
                str(idx0+1),
                list(sample_charges[LAB_TEST_CATEGORY_COL   ]),
                list(sample_charges[TEST_PARAMS_COL]),
                str(num_samples),
                sample[SAMPLE_LABELS_COL],
                sample[SAMPLE_DESCRIPTION_COL],
                "<TODO>"
            ]
        )

    return lab_test_details_data, lab_test_charge_data 

def create_docx_invoice(invoice, output_filename, labs_df):
    std_test_details_json = create_std_test_details_json(invoice, labs_df)
    # print(json.dumps(std_test_details_json, indent=4))
    print("*************************")
    std_test_charges_json = create_std_test_charges_json(invoice, labs_df)
    # print(json.dumps(std_test_charges_json, indent=4))
    shelf_life_details_json = create_shelf_life_test_details_json(invoice)
    shelf_life_charges_json = create_shelf_life_test_charges_json(invoice)
    total_lab_test_charges_json = create_total_lab_test_charges_json(shelf_life_charges_json, std_test_charges_json)
    food_service_charges_json = create_food_service_test_charges_json(invoice)
    total_cost_json = create_total_cost_json(invoice, total_lab_test_charges_json, food_service_charges_json)

    print(invoice)
    print(std_test_details_json)
    print(std_test_charges_json)
    print(shelf_life_details_json)
    print(shelf_life_charges_json)

    print(total_lab_test_charges_json)
    print(food_service_charges_json)
    print(total_cost_json)

    doc = Document()
    # add_doc_lib(doc)
    # add_doc_borders(doc)

    # test_param_details, test_charge_data = parse_json_input(test_details, labs_df)
    # print(f"Test Param details: {test_param_details}")
    # print(f"Test Charge details: {test_charge_data}")
                
    add_logo(doc)
    set_doc_font(doc)
    if std_test_details_json:
        add_std_test_details(doc, std_test_details_json)
    if shelf_life_details_json:
        add_shelf_life_test_details(doc, shelf_life_details_json)
    if std_test_charges_json:
        add_std_test_charges(doc, std_test_charges_json)
    if shelf_life_charges_json:
        add_shelf_life_test_charges(doc, shelf_life_charges_json)
    if total_lab_test_charges_json:
        add_total_lab_test_charges(doc, total_lab_test_charges_json)
    add_total_food_service_charges(doc, food_service_charges_json)
    add_total_cost(doc, total_cost_json)

    doc.save(output_filename)

# def create_docx_invoice_v2(test_details, output_filename, labs_df):
#     doc = Document()
#     add_doc_lib(doc)
#     add_doc_borders(doc)

#     test_param_details, test_charge_data = parse_json_input(test_details, labs_df)
#     print(f"Test Param details: {test_param_details}")
#     print(f"Test Charge details: {test_charge_data}")
                
#     add_logo(doc)
#     add_lab_test_details(doc, test_param_details)
#     add_shelflife_test_details(doc)
#     add_lab_test_charges(doc, test_charge_data)
#     add_shelflife_test_charges(doc)
#     add_total_lab_test_charges(doc)
#     add_total_food_service_charges(doc)
#     add_total_charges(doc)

#     doc.save(output_filename)

if __name__ == "__main__":

    #Input
    test_details = [
        {
            "sample_labels" : ["Protein Sample", "Protein Sample 2"],
            "sample_description" : "For Gym boys",
            "test_param_details" : [
                {
                    "test_param" : "Cystine", 
                    "lab_id" : "L030"
                },
                {
                    "test_param" : "Tryptophan", 
                    "lab_id" : "L026"
                }
            ]
        },
        {
            "sample_labels" : ["For some Acid Sample", "Acid Sample 2"],
            "sample_description" : "For making legal stuff",
            "test_param_details" : [
                {
                    "test_param" : "Sodium (Na)", 
                    "lab_id" : "L021"
                },
                {
                    "test_param" : "Phosphate (PO₄³⁻)", 
                    "lab_id" : "L021"
                }
            ]
        }
    ]

    df_dict = {}
    for sheet in ["Labs", "Test Packages", "Tests", "Test Params"]:
        df_dict[sheet] = pd.read_excel("Labs DB.xlsx", sheet_name=sheet)

    test_df = df_dict["Tests"]
    test_p_df = df_dict["Test Params"]
    test_df = test_df.merge(test_p_df, on="Test ID", how="outer")
    test_df[PRICE_COL] = test_df[PRICE_COL].fillna(0)
    test_df[BASE_PRICE_COL] = test_df[BASE_PRICE_COL].fillna(0)

    create_docx_invoice(test_details, output_filename, labs_df = test_df)

    # Todo in future : generate pdf
    # in_file = os.path.abspath(output_filename)
    # out_file = os.path.abspath(output_filename.replace(".docx", ".pdf"))
    # convert(in_file, out_file)