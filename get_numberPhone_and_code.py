import requests


def get_phone():
    url = "https://chothuesimcode.com/api?act=number&apik=3fbef40e&appId=1017&carrier=Viettel"
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    if response.json()["Msg"] == "OK":
        num_phone = response.json()["Result"]["Number"]
        id_phone = response.json()["Result"]["Id"]
        return num_phone, id_phone
    else:
        return -1, -1


def get_code(id_phone):

    url = f"https://chothuesimcode.com/api?act=code&apik=3fbef40e&id={id_phone}"
    print(url)
    payload = {}
    headers = {}
    response = requests.request("GET", url, headers=headers, data=payload)
    print("response", response.json())
    if response.json()["Msg"] == "Đã nhận được code":
        return response.json()["Result"]["Code"], response.json()["ResponseCode"]
    elif response.json()["Msg"] == "không nhận được code, quá thời gian chờ":
        return -2, -2
    else:
        return -1, -1

