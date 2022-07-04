from imap_tools import MailBox, AND
from datetime import datetime
from requests_html import HTML


def get_verify_code(mail: str, seen: bool):
    date = datetime.date(datetime.now())
    with MailBox('imap.yandex.com').login('nhanmailao@minh.live', 'Team123@') as mailbox:
        for msg in mailbox.fetch(criteria=AND(seen=seen, to=mail), reverse=True):
            page = HTML(html=str(msg.html))
            codes = page.xpath('//tr/td/span//text()')
            if len(codes):
                return codes[0]
            else:
                return -1


if __name__ == '__main__':
    mail_protect = "dentelstreb3@minh.live"
    code = get_verify_code(mail_protect, False)
    print(code)
