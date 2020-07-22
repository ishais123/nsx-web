import requests
import json
import xlsxwriter
from flask import Flask, render_template, redirect, url_for, request
import argparse

requests.packages.urllib3.disable_warnings()  # Ignore from requests module warnings


workbook = xlsxwriter.Workbook('dfw_policy.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A', 20)
bold = workbook.add_format({'bold': True})
worksheet.write(0, 0, "Rule Name")
worksheet.write(0, 1, "Source")
worksheet.write(0, 2, "Destenation")
worksheet.write(0, 3, "Service")
worksheet.write(0, 4, "Action")


def parse_args():
    parser = argparse.ArgumentParser(description='This is HELP doc for NSX-T Automation tool')
    parser.add_argument("-user", "--username", help="Your nsx username.")
    parser.add_argument("-pwd", "--password", help="Your nsx password.")
    parser.add_argument("-auth", "--authorization", help="Your nsx password.")
    args = parser.parse_args()
    return args


def get_sections_id():
    ids = []
    payload = {}
    headers = {
      'Authorization': 'Basic YWRtaW46QnBvdnRtZzEhQnBvdnRtZzEh',
      'Cookie': 'JSESSIONID=BCCFED32F7DFCB6233E1A6E545B00E90'
    }

    url = "https://172.16.20.103/policy/api/v1/infra/domains/default/security-policies"

    response = requests.request("GET", url, headers=headers, data=payload, verify=False)
    to_json = response.json()
    results = to_json.get('results')
    for result in results:
        id = result.get('id')
        ids.append(id)
    return ids


def split_groups_from_paths(groups_paths):
    groups = []
    for group_path in groups_paths:
        if group_path != "ANY":
            split_path = group_path.split('/')
            group = split_path[5]
        else:
            group = "ANY"
        groups.append(group)
    return groups


def split_services_from_paths(services_paths):
    services =[]
    for services_path in services_paths:
        split_path = services_path.split('/')
        service = split_path[3]
        services.append(service)
    return services


def nsx_api_call(section_id):
    payload = {}
    headers = {
      'Authorization': 'Basic YWRtaW46QnBvdnRtZzEhQnBvdnRtZzEh',
      'Cookie': 'JSESSIONID=BCCFED32F7DFCB6233E1A6E545B00E90'
    }

    url = f"https://172.16.20.103/policy/api/v1/infra/domains/default/security-policies/{section_id}/rules"

    response = requests.request("GET", url, headers=headers, data=payload, verify=False)
    return response


def build_excel():
    try:
        section_ids = get_sections_id()
        rule_count = 1
        for section_id in section_ids:
            response = nsx_api_call(section_id)
            to_json = response.json()
            rules = to_json.get('results')
            sources = []

            for rule in rules:
                name = rule.get('display_name')
                worksheet.write(rule_count, 0, name)

                source_paths = rule.get('source_groups')
                sources = split_groups_from_paths(source_paths)
                worksheet.write(rule_count, 1, str(sources))

                destanation_paths = rule.get('destination_groups')
                destanations = split_groups_from_paths(destanation_paths)
                worksheet.write(rule_count, 2, str(destanations))

                services_paths = rule.get('services')
                services = split_services_from_paths(services_paths)
                worksheet.write(rule_count, 3, str(services))

                action = rule.get('action')
                worksheet.write(rule_count, 4, action)

                rule_count = rule_count + 1

        workbook.close()
        return "DFW exported, please see 'dfw_pilocy.xlsx' doc in your folder"
    except:
        return "Failed to export DFW policy"


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        args = parse_args()
        nsx_username = args.username
        nsx_password = args.password
        if request.form['username'] != f'{nsx_username}' or request.form['password'] != f'{nsx_password}':
            error = 'Invalid Credentials. Please try again.'
        else:
            return redirect(url_for('main'))
    return render_template('login.html', error=error)


@app.route('/main', methods=['GET'])
def main():
    return render_template('main.html'), 201


@app.route('/run')
def run():
    return build_excel()


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=1111)
