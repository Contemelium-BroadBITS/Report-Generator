from hashlib import md5
from json import dumps, loads
from requests import Session
from urllib.parse import unquote
import os
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
import json
import re

class Interactions:
    
    def __init__(self):
        
        self.username = os.getenv('OPEXPERT_USERNAME')
        self.password = os.getenv('OPEXPERT_PASSWORD')
        self.CRMURL = os.getenv('OPEXPERT_CRM_URL')        

        self.sessionID = False

    def login(self):
        
        data = {
            'user_auth': {
                'user_name': self.username, 
                'password': self.password
            }
        }
        
        response = self.__call('login', data)
        self.sessionID = response.get('id') 

    def __call(self, method, data, URL = False):
        
        curlRequest = Session()
        
        payload = {
            'method': method, 
            'input_type': 'JSON', 
            'response_type': 'JSON', 
            'rest_data': dumps(data), 
            'script_command': True
        }
        
        response = curlRequest.post(URL if URL else self.CRMURL, data = payload)
        curlRequest.close()
        result = loads(response.text)

        return result

    def __createSession(self):
        self.session = Session()
        
    def __closeSession(self):
        try:
            self.session.close()
        except Exception:
            pass


    def getIntegrationWithID(self, reportID = '', params=''):
        if self.sessionID:
            print(params)
            data = {
                'session': self.sessionID, 
                'report_id': reportID, 
                'UserInputParams': base64.b64encode(params.encode()).decode()
            }
            try:
                return self.__call('getAPIReportResponse', data)
            except:
                return "An error occurred. Please try again after verifying your session ID and report ID."
            
        else:
            return "You cannot proceed with this action without initializing a session."
    
    
    def downloadDocumentWithID(self, document_id):
        crm_url = "https://apigw.uat.opexpert.io/custom/service/v4_1_custom/rest.php"
        session = requests.Session()

        # --- Step 1: Login ---
        payload = {
            "method": "login",
            "input_type": "JSON",
            "response_type": "JSON",
            "rest_data": dumps({
                "user_auth": {
                    "user_name": self.username,
                    "password": self.password
                }
            }),
            "script_command": True
        }
        resp = session.post(crm_url, data=payload)
        login_result = resp.json()

        if "id" not in login_result:
            raise Exception(f"Login failed: {login_result}")

        # --- Step 2: Download file ---
        # download_url = f"{crm_url}?entryPoint=download&id={document_id}&type=Documents"
        download_url = f"https://apigw.uat.opexpert.io/index.php?entryPoint=download&id={document_id}&type=Documents"
        r = session.get(download_url, allow_redirects=True)
        return r
    
    

    def getModuleWithID(self, reportID = '', moduleName = '', fields = []):
        
        if self.sessionID:
            
            data = {
                'session': self.sessionID, 
                'module_name': moduleName, 
                'query': f"{moduleName.lower()}.id = \'{reportID}\'", 
                'order_by': '', 
                'offset': 0, 
                'deleted': False
            }
            
            try:
                module = self.__call('get_entry_list', data)['entry_list']
                if len(fields) == 0:
                    return module
                elif len(fields) == 1:
                    return module[0]['name_value_list'][fields[0]]['value']
                else:
                    requiredFields = {}
                    for field in fields:
                        requiredFields[field] = module[0]['name_value_list'][field]['value']
                    return requiredFields
            except:
                return "An error occurred. Please try again after verifying your session ID and report ID."
            
        else:
            return "You cannot proceed with this action without initializing a session."

    def getCodeSnippetWithID(self, reportID = ''):
        
        if self.sessionID:
            
            data = {
                'session': self.sessionID, 
                'module_name': 'bc_api_methods', 
                'query': f"{'bc_api_methods'.lower()}.id = \'{reportID}\'", 
                'order_by': '', 
                'offset': 0, 
                'deleted': False
            }
            
            try:
                code = unquote(self.__call('get_entry_list', data)['entry_list'][0]['name_value_list']['description']['value'])
                return code if code else 'return None'
            except:
                return "An error occurred. Please try again after verifying your session ID and report ID."
            
        else:
            return "You cannot proceed with this action without initializing a session."
        
    def getHTMLTemplateWithID(self, reportID = ''):
        
        if self.sessionID:
            
            data = {
                'session': self.sessionID, 
                'module_name': 'bc_html_writer', 
                'id': f"{reportID}", 
                'select_fields': [], 
                'link_name_to_fields_array': [], 
                'track_view': False
            }
            
            try:
                code = unquote(self.__call('get_entry',data)['entry_list'][0]['name_value_list']['html_body']['value'])
                return code if code else 'return None'
            except:
                return "An error occurred. Please try again after verifying your session ID and report ID."
            
        else:
            return "You cannot proceed with this action without initializing a session."
        
    def getEmailTemplateWithID(self, reportID = ''):
        
        if self.sessionID:
            
            data = {
                'session': self.sessionID, 
                'module_name': "EmailTemplates", 
                'id': f"{reportID}", 
                'select_fields': [], 
                'link_name_to_fields_array': [], 
                'deleted': False
            }
            
            try:
                code = self.__call('get_entry',data)['entry_list'][0]['name_value_list']['body_html']['value']
                return code if code else 'return None'
            except:
                return "An error occurred. Please try again after verifying your session ID and report ID."
            
        else:
            return "You cannot proceed with this action without initializing a session."
        
    def getReport(yamlFile,emailConfig): 
        command = f"python3 -X tracemalloc=25 /home/rundeck/projects/RulesInterpreterApp02/getAPIReport.py {yamlFile} \"{emailConfig}\""
        os.system(command)

    @staticmethod
    def sendEmail(subject, body, recipients, variable_replacement):
        VAULT_ADDR = "https://vault.broadbits.com"
        ROLE_ID = "d999f7e8-b9ad-e2df-08c2-0b66fadeb1d8"
        SECRET_ID = "f683fa30-d8f5-0557-a30d-2dd5271f23d8"
        SECRET_PATH = "rules_engine/data/support"

        def get_vault_token(vault_addr, role_id, secret_id):
            login_url = f"{vault_addr}/v1/auth/approle/login"
            payload = {"role_id": role_id, "secret_id": secret_id}
            try:
                response = requests.post(login_url, json=payload)
                response_data = response.json()
                if "auth" in response_data and "client_token" in response_data["auth"]:
                    return response_data["auth"]["client_token"]
                else:
                    print("Authentication failed. No token received.")
                    return None
            except requests.RequestException as e:
                print(f"Error during authentication: {e}")
                return None

        def get_vault_secret(vault_addr, token, secret_path):
            secret_url = f"{vault_addr}/v1/{secret_path}"
            headers = {"X-Vault-Token": token}
            try:
                response = requests.get(secret_url, headers=headers)
                response_data = response.json()
                if "errors" in response_data:
                    print(f"Error retrieving the secret: {response_data['errors']}")
                    return None
                return response_data
            except requests.RequestException as e:
                print(f"Error retrieving the secret: {e}")
                return None
            
        def json_to_html_table(data):
            if isinstance(data, str):
                try:
                    data = json.loads(data)
                except json.JSONDecodeError:
                    return f"<p>{data}</p>"
            
            if not data:
                return "<p>No data available.</p>"
                
            if isinstance(data, dict):
                data = [data]
                
            if not isinstance(data, list):
                return f"<p>{str(data)}</p>"
                
            headers = sorted({k for item in data for k in item.keys()})
            
            html = '<table style="border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 10px;">'
            # Table header
            html += '<thead><tr style="background-color: #f2f2f2;">'
            for header in headers:
                html += f'<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">{header}</th>'
            html += '</tr></thead>'
            
            # Table body
            html += '<tbody>'
            for i, item in enumerate(data):
                row_style = 'background-color: #f9f9f9;' if i % 2 == 0 else ''
                html += f'<tr style="{row_style}">'
                for header in headers:
                    value = item.get(header, "")
                    html += f'<td style="border: 1px solid #ddd; padding: 8px;">{value}</td>'
                html += '</tr>'
            html += '</tbody></table>'
            
            return html

        def create_html_template(content):
            return f'''
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>{subject}</title>
            </head>
            <body style="font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 0;">
                <div style="background-color: #1e88e5; color: white; padding: 20px; text-align: center;">
                    <h1 style="margin: 0;">OpExpert Notification</h1>
                </div>
                <div style="padding: 20px; max-width: 800px; margin: 0 auto;">
                    {content}
                </div>
                <div style="background-color: #f5f5f5; padding: 10px; text-align: center; font-size: 12px; color: #666;">
                    <p>This is an automated message from OpExpert. Please do not reply to this email.</p>
                </div>
            </body>
            </html>
            '''

        # Dynamically replace all variable placeholders in the body using variable_replacement
        def replace_variables_in_body(body, variable_replacement):
            # Find all placeholders like {var}
            placeholders = re.findall(r'\{(\w+)\}', body)
            for var_name in placeholders:
                if var_name in variable_replacement:
                    var_value = variable_replacement[var_name]
                    # If value is list/dict, format as table
                    try:
                        parsed = json.loads(var_value) if isinstance(var_value, str) else var_value
                        if isinstance(parsed, (list, dict)):
                            table = json_to_html_table(parsed)
                            body = body.replace(f'{{{var_name}}}', table)
                        else:
                            body = body.replace(f'{{{var_name}}}', str(parsed))
                    except Exception:
                        body = body.replace(f'{{{var_name}}}', str(var_value))
            return body

        # Replace variables and convert to HTML
        html_body = replace_variables_in_body(body, variable_replacement)
        # Wrap the content in our HTML template
        html_body = create_html_template(html_body)
        
        print("Sending email with HTML content")

        token = get_vault_token(VAULT_ADDR, ROLE_ID, SECRET_ID)
        if not token:
            print("Could not retrieve Vault token.")
            return False
        secret_response = get_vault_secret(VAULT_ADDR, token, SECRET_PATH)
        if not secret_response:
            print("Could not retrieve Vault secret.")
            return False

        data = secret_response['data']['data']
        username = data.get('username')
        password = data.get('password')
        smtp_server = data.get('smtp_server')
        port = data.get('port')

        message = MIMEMultipart('alternative')
        message['From'] = f'OpExpert Notifications <{username}>'
        message['To'] = ', '.join(recipients)
        message['Subject'] = subject
        
        # Attach plain text and HTML versions
        message.attach(MIMEText(body, "plain"))
        message.attach(MIMEText(html_body, "html"))

        try:
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()
            server.login(username, password)
            server.sendmail(username, recipients, message.as_string())
            server.quit()
            print(f"Alert email sent successfully to {recipients}")
            return True
        except Exception as e:
            print(f"Failed to send alert email: {e}")
            return False



if __name__ == '__main__':
    
    an_object = Interactions()
    an_object.login()
    response = an_object.downloadDocumentWithID('c7f6af3c-09e0-1e14-b788-68b4ee963318', '/datadrive/Afraaz/Generator/20250827/template.docx')
    # if hasattr(response, 'status_code') and response.status_code == 200:
    #     out_path = '/datadrive/Afraaz/Generator/20250827/template.docx'
    #     with open(out_path, 'wb') as f:
    #         for chunk in response.iter_content(chunk_size=8192):
    #             if chunk:
    #                 f.write(chunk)
    #     print('Saved to', out_path)
    # else:
    #     print('Download failed:', getattr(response, 'status_code', response))
    # print(response.content)
    # print(an_object.getIntegrationWithID('6adaaae2-5e26-4b94-4ea2-64a8268fe518'))
