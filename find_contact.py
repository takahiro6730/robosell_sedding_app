import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

class FindContact:
    @staticmethod
    def find_contact_page_url(base_url, target_texts=['お問合せ', 'お問い合わせ', 'CONTACT', 'CONTACT US', 'ご相談', 'contact','Contact','企業様向け相談','相談ボックス','相談フォーム'], timeout=10):
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
            response = requests.get(base_url, headers=headers, allow_redirects=True, timeout=timeout)

            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                for target_text in target_texts:
                    contact_links = soup.find_all('a')
                    contact_links = [link for link in contact_links if link.has_attr('href') and target_text in link.get_text()]
                    if len(contact_links) > 0:
                        contact_page_url = urljoin(base_url, contact_links[0]['href'])
                        response = requests.get(contact_page_url, headers=headers, allow_redirects=True, timeout=timeout)
                        soup = BeautifulSoup(response.content, 'html.parser')
                        form_tag = soup.find('form')

                        if form_tag:
                            return contact_page_url
                        else:
                            contact_links = soup.find_all('a')
                            contact_links = [link for link in contact_links if link.has_attr('href') and urljoin(base_url, link['href']) != contact_page_url and target_text in link.get_text()]

                            if len(contact_links) == 1:
                                contact_page_url = urljoin(contact_page_url, contact_links[0]['href'])
                            elif len(contact_links) > 1:
                                contact_page_url = urljoin(contact_page_url, contact_links[-1]['href'])
                            response = requests.get(contact_page_url, headers=headers, allow_redirects=True, timeout=timeout)
                            soup = BeautifulSoup(response.content, 'html.parser')
                            form_tag = soup.find('form')
                            if form_tag:
                                return contact_page_url
                            else:
                                return False

                return False
            else:
                return False

        except requests.RequestException as e:
            return False
        
    ALLOWED_INPUT_TYPES = ['text', 'url', 'email', 'tel']
    EXCLUDED_LABEL_TEXT = [' ','-', '必須','(必須)', '*', '任意', '（※半角数字でご記入ください）', '※確認のため、再度入力してください', '（※カタカナでご記入ください）', '（※漢字でご記入ください）', '※', '法人の場合必須', '-  -', '（※半角英数字でご記入ください）', '--', '（必須）']

    def find_contact_form(self, url):
        try:
            response = requests.get(url)
            response.raise_for_status()

            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                keywords = ["氏名", "電話番号", "メールアドレス", "お名前", "住所", "お問合せ", "ご連絡先", "部署名", "Email", "Tel", "住所", "お問合せ内容"]
                
                target_form = None
                for form in soup.find_all('form'):
                    form_text = form.get_text()
                    if any(keyword in form_text for keyword in keywords):
                        input_fields = form.find_all('input')
                        if input_fields:
                            target_form = form
                            break

                if target_form:
                    input_labels = self.extract_input_labels(target_form)
                    return(input_labels)
                else:
                    print(f"No contact form found with the specified keywords.")
                    return {}
            else:
                print(f"Failed to retrieve the webpage. Status code: {response.status_code}")
                return {}
        except requests.exceptions.RequestException as e:
            print(f"Failed to retrieve the webpage. Error: {e}")
            return {}
        
    def extract_input_labels(self, form):
        input_labels = {}

        input_fields = form.find_all('input', type=lambda x: x and x.lower() in self.ALLOWED_INPUT_TYPES)

        for input_field in input_fields:
            if input_field.find_parents(['select', 'option']):
                continue

            label_text = self.get_label_text(input_field, input_labels)
            input_name = input_field.get('name', '')
            input_type = input_field.get('type', '')
            input_value = input_field.get('value', '')

            if label_text:
                if label_text in input_labels:
                    input_labels[label_text].append({'name': input_name, 'type': input_type, 'value': input_value})
                else:
                    input_labels[label_text] = [{'name': input_name, 'type': input_type, 'value': input_value}]

        return input_labels

    def get_label_text(self, input_field, input_labels):
        label_text = None
        current_text = input_field.find_previous(['label', 'p', 'th', 'dt', 'span', 'div', 'td']).text.strip()
        current_th = input_field.find_previous('th')
        current_label = input_field.find_previous('label')


        if current_text and current_text not in self.EXCLUDED_LABEL_TEXT and '例' not in current_text and 'ください' not in current_text:
            label_text = current_text
        else:
            previous_element = input_field.find_previous(['label', 'p', 'th', 'dt', 'span', 'div', 'td'])
            while previous_element and not label_text:
                previous_text = previous_element.text.strip()
                if previous_text and previous_text not in self.EXCLUDED_LABEL_TEXT and '例' not in previous_text and 'ください' not in previous_text:
                    label_text = previous_text
                else:
                    previous_element = previous_element.find_previous(['label', 'p', 'th', 'dt', 'span', 'div', 'td'])

        if current_th and '電話番号' in current_th.get_text():
            label_text = '電話番号'
        if current_label and '電話番号' in current_label.get_text():
            label_text = '電話番号'

        return label_text