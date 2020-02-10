import requests

TIMEOUT = 10
HEADERS = {
    "User-Agent":
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) \
         AppleWebKit/537.36 (KHTML, like Gecko) \
         Chrome/59.0.3071.115 Safari/537.36",
}


def getResponse(url):
    try:
        resp = requests.get(
            url,
            headers=HEADERS,
            timeout=TIMEOUT,
        )

        return resp
    except Exception as e:
        return None


def getM3u8File(url):
    m3u8_content = getResponse(url).text
    m3u8_lines = m3u8_content.split('\n')
    # print(m3u8_lines)
    return m3u8_lines
