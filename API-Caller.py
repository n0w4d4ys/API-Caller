import argparse
import sys
import xlwt
import xlrd
import requests
from xlutils.copy import copy as xlcopy


def read_book(path):
    xlsx_file = xlrd.open_workbook(path, on_demand=True, encoding_override="utf8")
    return xlsx_file, xlsx_file.sheet_by_index(0), xlsx_file.sheet_by_index(1)


def edit_book(xlsx_file):
    draft_xlsx = xlcopy(xlsx_file)
    return draft_xlsx, draft_xlsx.get_sheet(0), draft_xlsx.get_sheet(1)


def save_book(xlsx_file, path):
    xlsx_file.save(path)
    return None


def read_config_from_file(sheet):
    """ Get the first page of the XLSX file and returns URL and Headers to make requests """
    # Check URL is valid
    ping = requests.get(sheet.cell_value(0, 1), headers=eval(sheet.cell_value(1, 1)))
    ping.raise_for_status()
    return sheet.cell_value(0, 1), sheet.cell_value(1, 1)


def read_test_cases_from_file(sheet):
    """ Get the second page of the XLSX file and read a list of test cases in the following format: METHOD, POSTFIX,
    BODY, EXPECTED STATUS CODE, EXPECTED SERVER RESPONSE
    returns sheet, numbers of columns and rows
    """
    return sheet, sheet.ncols, sheet.nrows


def make_request(method, url, headers, body):
    """ Supported methods: GET, POST, PUT, PATCH, DELETE """
    if method.lower().strip() == "get":
        req = requests.get(url, headers=eval(headers))
    elif method.lower().strip() == "post":
        req = requests.post(url, headers=eval(headers), data=body)
    elif method.lower().strip() == "put":
        req = requests.put(url, headers=eval(headers), data=body)
    elif method.lower().strip() == "patch":
        req = requests.patch(url, headers=eval(headers), data=body)
    elif method.lower().strip() == "delete":
        req = requests.delete(url, headers=eval(headers))
    else:
        raise Exception("Unknown REST API method")
    return req


def results_writer(sheet, rw, act_code, act_res, act_head, res, note):
    sheet.write(rw, 5, act_code)
    sheet.write(rw, 6, act_res)
    sheet.write(rw, 7, str(act_head))
    if res == "PASS":
        style = xlwt.easyxf('font: bold 1, color-index green')
    elif res == "FAIL":
        style = xlwt.easyxf('font: bold 1, color-index red')
    else:
        style = xlwt.easyxf('font: bold 1, color-index yellow')
    sheet.write(rw, 8, res, style)
    style = xlwt.easyxf('font: bold 1')
    sheet.write(rw, 9, note, style)


def marker(path):
    number_of_failed_tests = 0
    book = read_book(path)[2]
    rows_num = read_test_cases_from_file(book)[2]
    for i in range(1, rows_num):
        if book.cell_value(i, 8) == "FAIL":
            number_of_failed_tests = number_of_failed_tests + 1
        else:
            pass
    if number_of_failed_tests == 0:
        pass
    else:
        raise Exception("We have %s failed test cases" % number_of_failed_tests)


def compare_status_codes(expected_result, actual_result):
    if expected_result in (None, ""):
        result = ""
        note = ""
    else:
        if actual_result != expected_result:
            result = "FAIL"
            note = "Wrong status code."
        else:
            result = "PASS"
            note = ""
    return result, note


def compare_server_responses(expected_result, actual_result):
    if expected_result in (None, ""):
        result = ""
        note = ""
    else:
        # Normalize
        expected_result = expected_result.strip()
        actual_result = actual_result[:32766].strip()

        if actual_result != expected_result:
            result = "FAIL"
            note = "Wrong server response."
        else:
            result = "PASS"
            note = ""
    return result, note


def tests_generator(path):
    # Read a book
    xlsx = read_book(path)
    # Make a draft of book
    xlsx_draft = edit_book(xlsx[0])
    # Get URL and Headers:
    configs = read_config_from_file(xlsx[1])
    # Get a set of test cases
    set_of_cases = read_test_cases_from_file(xlsx[2])
    # Perform tests
    if set_of_cases[2] <= 1 or set_of_cases[1] <= 1:
        raise Exception("The provided file doesn't contain test cases")
    else:
        for i in range(1, set_of_cases[2]):
            test_method = set_of_cases[0].cell_value(i, 0)
            test_postfix = set_of_cases[0].cell_value(i, 1).strip()
            test_body = set_of_cases[0].cell_value(i, 2)
            test_ex_status_code = set_of_cases[0].cell_value(i, 3)
            test_exp_ser_response = set_of_cases[0].cell_value(i, 4)

            response = make_request(test_method,
                                    configs[0]+test_postfix,
                                    configs[1],
                                    test_body)

            status_codes_result = compare_status_codes(test_ex_status_code, response.status_code)
            server_responses_result = compare_server_responses(test_exp_ser_response, response.text)

            if "FAIL" in (status_codes_result[0], server_responses_result[0]):
                res = "FAIL"
            else:
                res = "PASS"
            note = status_codes_result[1] + server_responses_result[1]

            results_writer(xlsx_draft[2], i, response.status_code, response.text[:32766], response.headers, res, note)

    save_book(xlsx_draft[0], path)
    marker(path)


if __name__ == '__main__':
    path_to_file = ""
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser()
        parser.add_argument("-p", "--path")
        args, unknown = parser.parse_known_args()
        path_to_file = args.path
        if path_to_file in (None, ""):
            pass
        else:
            tests_generator(path_to_file)
    else:
        tests_generator(path_to_file)
