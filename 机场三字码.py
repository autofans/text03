import requests
import xlsxwriter
from lxml import etree
import time


class Spider(object):

    def __init__(self):

        self.begin = begin

        self.end = end

        self.c = []

        self.t = []

        self.f = []

        self.a = []

        self.en = []

        self.header = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36"}

    def web_spider(self):
        """构建每一页的URL用xpath提取数据"""

        for i in range(self.begin, self.end+1):

            url = "https://airport.supfree.net/index.asp?page=%d" % i

            response = requests.get(url, headers=self.header)

            response.encoding = "gb18030"

            html = response.text

            # print(html)

            data = etree.HTML(html)

            all_tr = data.xpath('//div/table/tr')[1:]        # 用xpath抓取除第一个以外的所有tr

            for x in all_tr:
                city = x.xpath('.//td[1]//text()')           # 遍历城市名
                self.c.extend(city)

                if len(x.xpath('.//td[2]//text()')) > 0:     # 判断td里面的字符长度是否是空值
                    three = x.xpath('.//td[2]//text()')      # 抓取机场三字码
                else:
                    three = [None]                           # 空值的话用None填充

                if len(x.xpath('.//td[3]//text()')) > 0:
                    four = x.xpath('.//td[3]//text()')       # 抓取机场四字码
                else:
                    four = [None]

                if len(x.xpath('.//td[4]//text()')) > 0:
                    airport = x.xpath('.//td[4]//text()')    # 抓取机场名
                else:
                    airport = [None]

                if len(x.xpath('.//td[5]//text()')) > 0:
                    en_name = x.xpath('.//td[5]//text()')    # 抓取英文名
                else:
                    en_name = [None]

                self.t.extend(three)
                self.f.extend(four)
                self.a.extend(airport)
                self.en.extend(en_name)

            print(self.c)
            print(self.t)
            print(self.f)
            print(self.a)
            print(self.en)

            self.excel_spider()

    def excel_spider(self):
        """构造写入excel表格的函数"""

        workbook = xlsxwriter.Workbook("demo2.xlsx")

        worksheet = workbook.add_worksheet()

        for h in range(len(self.c)):
            worksheet.write("A%d" % int(h+1), self.c[h])

        for k in range(len(self.t)):
            worksheet.write("B%d" % int(k+1), self.t[k])

        for j in range(len(self.f)):
            worksheet.write("C%d" % int(j+1), self.f[j])

        for m in range(len(self.a)):
            worksheet.write("D%d" % int(m+1), self.a[m])

        for n in range(len(self.en)):
            worksheet.write("E%d" % int(n+1), self.en[n])

        workbook.close()

        time.sleep(1)


if __name__ == "__main__":
    """执行的主程序"""

    begin = int(input("请输入要爬取的起始页："))

    end = int(input("请输入要爬取的结束页："))

    my_spider = Spider()

    my_spider.web_spider()
