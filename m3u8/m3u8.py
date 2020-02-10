from util import getResponse, getM3u8File


class M3U8(object):

    def __init__(self, url):

        self.encrypt_method = None
        self.key_uri = None
        self.encrypt_iv = None
        self.ts_urls = []
        self.base_url = url.rsplit('/', 1)[0]

        self.m3u8_lines = self.parseM3u8Url(url)

        if self.m3u8_lines:
            self.parseTsUrl(self.base_url, self.m3u8_lines)
        else:
            print('Parse m3u8 url error!')

    def parseM3u8Url(self, url):
        # 解析M3U8 url，判断是否存在跳转
        m3u8_contents = None
        while m3u8_contents is None:
            m3u8_contents = getResponse(url).text

        if 'EXT-X-STREAM-INF' in m3u8_contents:
            # 存在跳转，需重新组合真实路径
            m3u8_lines = m3u8_contents.split('\n')
            actual_url = None
            for index, line_content in enumerate(m3u8_lines):
                if '.m3u8' in line_content:
                    actual_url = line_content
                    break

            url = self.base_url + '/' + actual_url
            print('Use embeded URL:%s' % url)
        return getM3u8File(url)

    def parseTsUrl(self, base_url, m3u8_lines):
        for index, line_content in enumerate(m3u8_lines):
            if 'EXT-X-KEY' in line_content:
                # 解析密钥
                content_units = line_content.split(',')
                for content_unit in content_units:
                    if 'METHOD' in content_unit:
                        if self.encrypt_method is None:
                            self.encrypt_method = content_unit.split('=')[-1]

                    elif 'URI' in content_unit:
                        if 'ccb.com' in content_unit:
                            self.key_uri = content_unit.split('"')[1]
                        else:
                            key_path = content_unit.split('"')[1]

                            self.key_uri = base_url + '/' + key_path  # 拼出key解密密钥URL
                    elif 'IV' in content_unit:
                        if self.encrypt_iv is None:
                            self.encrypt_iv = content_unit.split('=')[-1]
            if 'EXTINF' in line_content:
                # 拼出ts片段的URL
                if m3u8_lines[index + 1].startswith('/'):
                    ts_url = base_url + m3u8_lines[index + 1]
                else:
                    ts_url = base_url + '/' + m3u8_lines[index + 1]
                self.ts_urls.append(ts_url)
