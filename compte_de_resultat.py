#!/usr/bin/python
# _*_ coding: utf-8

import calendar
import collections
import datetime
import json
import operator
import re
import webbrowser
from time import sleep
from urllib import quote, unquote

import click
import xlsxwriter
from dateutil.relativedelta import relativedelta
from exactonline.api import ExactApi
from exactonline.resource import GET
from exactonline.storage import IniStorage


class MyIniStorage(IniStorage):
    def get_response_url(self):
        "Configure your custom response URL."
        return self.get_base_url()


@click.group()
def cli():
    pass


@click.command("setup")
@click.option("--base-url", prompt="Base url", help="Base url")
@click.option(
    "--client-id", prompt="Client Id", help="Client id, found in the app center"
)
@click.option(
    "--client-secret",
    prompt="Client secret",
    help="Client secret, found in the app center",
)
def setup(base_url, client_id, client_secret):
    """One time setup of the ini storage. This onlu need to run once"""
    storage = get_storage()
    storage.set("application", "base_url", base_url)
    storage.set("application", "client_id", client_id)
    storage.set("application", "client_secret", client_secret)
    api = ExactApi(storage)
    # Open a web browser to get the code query parameter
    print(
        "A web browser is going to open in 5 seconds, authenticate and copy and paste the value of the code query parameter"
    )
    sleep(5)
    webbrowser.open_new(api.create_auth_request_url())
    code = unquote(raw_input("Enter value of the code query parameter: "))
    api.request_token(code)
    print("High five!! it worked")


@click.command("excel")
@click.option("--annee_fiscale", help="Année fiscale", type=int)
def excel(annee_fiscale):
    storage = get_storage()
    api = ExactApi(storage)
    # divisions, current_divisions = api.get_divisions()
    current_divisions = 17923
    print("Current division {}".format(current_divisions))

    all_accounts = api.rest(
        GET(
            "v1/%d/financial/GLAccounts?$filter=startswith(Code,'6')+eq+true+or+startswith(Code,'7')+eq+true"
            % current_divisions
        )
    )

    result = collections.OrderedDict()
    tableau = {}
    with open("rapport_config.json") as data_file:
        tableau = json.load(data_file, object_pairs_hook=collections.OrderedDict)

    operator_map = {"+": operator.iadd, "-": operator.isub}

    now = datetime.datetime.now()
    current_year = now.year if now.month < 9 else now.year + 1
    current_year = annee_fiscale if annee_fiscale else current_year

    starting_date = datetime.date(current_year - 1, 10, 1)
    end_date = datetime.date(current_year, 9, 30)

    number_of_months = (
        (end_date.year - starting_date.year) * 12
        + end_date.month
        - starting_date.month
        + 1
    )

    months = range(number_of_months)

    used_accounts = []
    for key, value in tableau.iteritems():
        if is_subcategory(value):
            subcategories = value
            category_name = key
            result[category_name] = {
                "total": [0] * number_of_months,
                "detail": collections.OrderedDict(),
                "type": "subcategory",
            }
            for subcategory_name, account_list in subcategories.iteritems():
                result[category_name]["detail"][subcategory_name] = {
                    "total": [0] * number_of_months,
                    "detail": collections.OrderedDict(),
                    "type": "account_list",
                }

                for (_, account_number) in account_list:
                    # account_name = find_account_name(account_number, all_accounts)
                    result[category_name]["detail"][subcategory_name]["detail"][
                        account_number
                    ] = [0] * number_of_months
                    used_accounts.append(account_number)

        elif is_account_list(value):
            category_name = key
            account_list = value
            result[category_name] = {
                "total": [0] * number_of_months,
                "detail": collections.OrderedDict(),
                "type": "account_list",
            }

            for (_, account_number) in account_list:
                # account_name = find_account_name(account_number, all_accounts)
                result[category_name]["detail"][account_number] = [0] * number_of_months
                used_accounts.append(account_number)

        elif is_total_list(value):
            total_name = key
            category_list = value
            result[total_name] = {"total": [0] * number_of_months, "type": "total_list"}

    unused_lines = []
    for month_index in months:
        report_date = starting_date + relativedelta(months=month_index)
        financial_lines = request_financial_lines(
            api, current_divisions, report_date.year, report_date.month
        )
        unused_lines = unused_lines + find_unused_lines(financial_lines, used_accounts)

        for key, value in tableau.iteritems():
            if is_subcategory(value):
                subcategories = value
                category_name = key
                for subcategory_name, account_list in subcategories.iteritems():
                    for (sign, account_number) in account_list:
                        make_account_list_totals(
                            operator_map[sign],
                            account_number,
                            financial_lines,
                            month_index,
                            result[category_name]["detail"][subcategory_name]["detail"][
                                account_number
                            ],
                            result[category_name]["detail"][subcategory_name]["total"],
                            result[category_name]["total"],
                        )
            elif is_account_list(value):
                category_name = key
                account_list = value
                for (sign, account_number) in account_list:
                    make_account_list_totals(
                        operator_map[sign],
                        account_number,
                        financial_lines,
                        month_index,
                        result[category_name]["total"],
                        result[category_name]["detail"][account_number],
                    )
            elif is_total_list(value):
                total_name = key
                category_list = value
                for (sign, category_name) in category_list:
                    make_total_list_totals(
                        operator_map[sign],
                        category_name,
                        category_list,
                        result,
                        month_index,
                        result[total_name]["total"],
                    )

    unused_accounts = [
        account for account in all_accounts if int(account["Code"]) not in used_accounts
    ]

    workbook = xlsxwriter.Workbook(
        "compte_de_resultat-{}.xlsx".format(now.strftime("%Y%m%d%H%M%S"))
    )
    worksheet = workbook.add_worksheet("Rapport")
    worksheet.fit_to_pages(1, 0)  # 1 page wide and as long as necessary.
    worksheet.set_landscape()
    worksheet.repeat_rows(0, 2)

    worksheet.set_row(0, 30)
    date_header_format = workbook.add_format({"align": "center"})
    title_format = workbook.add_format({"bold": 1, "border": 1, "align": "center"})
    title_format.set_align("vcenter")
    category_format = workbook.add_format({"bold": 1})
    subcategory_format = workbook.add_format({"bold": 1})
    money_format = workbook.add_format({"num_format": "#,##0.00"})
    total_format = workbook.add_format({"bold": 1})

    worksheet.merge_range(
        "A1:N1",
        u"Compte de resultat SIF du %s au %s"
        % (starting_date.strftime("%m/%Y"), end_date.strftime("%m/%Y")),
        title_format,
    )

    worksheet.write(2, 1, u"Intitulé de la ligne"),

    [
        worksheet.write(
            2,
            2 + month_index,
            (starting_date + relativedelta(months=month_index)).strftime("%m/%Y"),
            date_header_format,
        )
        for month_index in months
    ]

    worksheet.set_column(1, 1, 35)
    worksheet.set_column(2, 13, 12)

    row = 4
    col = 0

    for category in result:
        if result[category]["type"] == "subcategory":
            worksheet.write(row, col + 1, category, category_format)
            row += 1
            for subcategory in result[category]["detail"]:
                worksheet.write(
                    row, col + 1, u"{}".format(subcategory), subcategory_format
                )
                row += 1
                for account_number in result[category]["detail"][subcategory]["detail"]:
                    account_description = find_account_name(
                        account_number, all_accounts
                    )
                    worksheet.write(row, col, account_number)
                    worksheet.write(
                        row,
                        col + 1,
                        u"{}".format(account_description).lower().capitalize(),
                    )
                    account_total = result[category]["detail"][subcategory]["detail"][
                        account_number
                    ]
                    [
                        worksheet.write(row, col + 2 + index, total, money_format)
                        for index, total in enumerate(account_total)
                    ]
                    row += 1
                subcategory_total = result[category]["detail"][subcategory]["total"]
                worksheet.write(
                    row, col + 1, u"Total %s" % subcategory, subcategory_format
                )
                [
                    worksheet.write(row, col + 2 + index, total, money_format)
                    for index, total in enumerate(subcategory_total)
                ]
                row += 2
            category_total = result[category]["total"]
            worksheet.write(row, col + 1, u"TOTAL %s" % category)
            [
                worksheet.write(row, col + 2 + index, total, money_format)
                for index, total in enumerate(category_total)
            ]
            row += 2
        elif result[category]["type"] == "account_list":
            worksheet.write(row, col + 1, category, category_format)
            row += 1
            for account_number in result[category]["detail"]:
                value = result[category]["detail"][account_number]
                account_description = find_account_name(account_number, all_accounts)
                worksheet.write(row, col, account_number)
                worksheet.write(
                    row, col + 1, u"{}".format(account_description.lower().capitalize())
                )
                [
                    worksheet.write(row, col + 2 + index, total, money_format)
                    for index, total in enumerate(value)
                ]
                row += 1
            category_total = result[category]["total"]
            worksheet.write(row, col + 1, u"TOTAL %s" % category)
            [
                worksheet.write(row, col + 2 + index, total, money_format)
                for index, total in enumerate(category_total)
            ]
            row += 2
        elif result[category]["type"] == "total_list":
            worksheet.write(row, col + 1, u"{}".format(category))
            total = result[category]["total"]
            [
                worksheet.write(row, col + 2 + index, total, money_format)
                for index, total in enumerate(total)
            ]
            row += 2

    unused_lines_worksheet = workbook.add_worksheet(
        u"lignes non utilisées ({})".format(len(unused_lines))
    )
    row = 0
    col = 0
    unused_lines_worksheet.write_row(
        row,
        col,
        (
            u"Date",
            u"Numéro de compte",
            u"Description du compte",
            u"Numéro d'écriture",
            u"Montant",
        ),
    )
    row += 1
    [
        unused_lines_worksheet.write_row(
            row + index,
            col,
            (
                parse_date(line["Date"]).strftime("%d/%m/%Y"),
                line["GLAccountCode"],
                u"{}".format(line["GLAccountDescription"]),
                line["EntryNumber"],
                line["AmountDC"],
            ),
        )
        for index, line in enumerate(unused_lines)
    ]

    unused_accounts_worksheet = workbook.add_worksheet(
        u"Comptes non utilisés ({})".format(len(unused_accounts))
    )
    row = 0
    col = 0
    unused_accounts_worksheet.write(row, col, u"Numéro de compte")
    unused_accounts_worksheet.write(row, col + 1, u"Description du compte")

    row += 1
    [
        unused_accounts_worksheet.write_row(
            row + index, col, (account["Code"], u"{}".format(account["Description"]))
        )
        for index, account in enumerate(unused_accounts)
    ]

    workbook.close()


def get_storage():
    return MyIniStorage("config.ini")


def parse_date(date):
    return datetime.datetime.fromtimestamp(
        float(re.search("/Date\((.*)\)/", date).group(1)) / 1000
    )


"""
A category has a sub category if is a list of sub_category
"""


def is_subcategory(category):
    return type(category) is collections.OrderedDict


def is_account_list(category):
    return type(category[0][1]) is int


def is_total_list(category):
    return type(category[0][1]) is unicode


def find_unused_lines(financial_lines, used_accounts):
    return [
        line
        for line in financial_lines
        if int(line["GLAccountCode"]) not in used_accounts
        and (
            line["GLAccountCode"].startswith("6")
            or line["GLAccountCode"].startswith("7")
        )
    ]


def find_account_name(account_number, all_accounts):
    account_name = next(
        (
            account["Description"]
            for account in all_accounts
            if account["Code"] == str(account_number)
        ),
        str(account_number),
    )
    return account_name


def find_account_total(account_number, all_financial_lines):
    account_total = sum(
        line["AmountDC"]
        for line in all_financial_lines
        if line["GLAccountCode"] == str(account_number)
    )
    return account_total


def request_financial_lines(api, current_division, year, month):
    first_of_the_month = datetime.date(year, month, 1).strftime("%Y-%m-%d")
    last_of_the_month = datetime.date(
        year, month, calendar.monthrange(year, month)[1]
    ).strftime("%Y-%m-%d")
    date_filter = "Date+ge+DateTime'%s'+and+Date+le+DateTime'%s'" % (
        first_of_the_month,
        last_of_the_month,
    )
    request = (
        "v1/%d/bulk/Financial/TransactionLines?$select=AmountDC,Date,EntryNumber,FinancialPeriod,FinancialYear,GLAccountCode,GLAccountDescription&$filter=%s"
        % (current_division, date_filter)
    )
    return api.rest(GET(request))


def make_account_list_totals(
    operation, account_number, financial_lines, month_index, *accumulateurs
):
    account_total = find_account_total(account_number, financial_lines)
    for accumulateur in accumulateurs:
        accumulateur[month_index] = operation(accumulateur[month_index], account_total)


def make_total_list_totals(
    operation, category_name, category_list, result, month_index, *accumulateurs
):
    category_total = next(
        (
            category["total"][month_index]
            for (category_label, category) in result.iteritems()
            if category_label == category_name
        ),
        0,
    )
    for accumulateur in accumulateurs:
        accumulateur[month_index] = operation(accumulateur[month_index], category_total)


cli.add_command(excel)
cli.add_command(setup)

if __name__ == "__main__":
    cli()
