import requests


def request_img_as_bytes(url):
    r = requests.get(url)
    r.raise_for_status()
    return r.content