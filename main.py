import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By


def setup_browser():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_argument('--disable-blink-features=AutomationControlled')
    browser = webdriver.Chrome(options=options)

    browser.execute_cdp_cmd(
        'Page.addScriptToEvaluateOnNewDocument',
        {'source': "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"}
    )

    browser.get("https://kns.cnki.net/kns8s/AdvSearch")
    time.sleep(1)
    return browser


def search_article(browser, journal, title):
    browser.find_element(By.XPATH, "//*[@id='gradetxt']/dd[1]/div[2]/input").send_keys(title)
    browser.find_element(By.XPATH, "//*[@id='gradetxt']/dd[3]/div[2]/input").send_keys(journal)
    browser.find_element(By.XPATH, "//div[@class='search-buttons']/input").click()
    time.sleep(2)


def open_article_detail(browser):
    try:
        href = browser.find_element(By.XPATH, "//tr[1]//td[@class='name']/a").get_attribute('href')
        browser.execute_script('window.open()')
        browser.switch_to.window(browser.window_handles[1])
        browser.get(href)
        time.sleep(1)
        return True
    except:
        return False


def extract_and_save_paper(browser, source_author, output_file):
    try:
        title = browser.find_element(By.XPATH, '//div[@class="wx-tit"]//h1').text
        author = browser.find_element(By.XPATH, '//div[@class="wx-tit"]//h3[1]').text
        institute = browser.find_element(By.XPATH, '//div[@class="wx-tit"]//h3[2]').text
        journal = browser.find_element(By.XPATH, "//div[@class='top-tip']/span/a[1]").text
        issue = browser.find_element(By.XPATH, "//div[@class='top-tip']/span/a[2]").text
        page_num = browser.find_element(By.XPATH, '//p[@class="total-inform"]/span[1]').text
        page_number = browser.find_element(By.XPATH, '//p[@class="total-inform"]/span[2]').text
        page_count = browser.find_element(By.XPATH, '//p[@class="total-inform"]/span[3]').text
        size = browser.find_element(By.XPATH, '//p[@class="total-inform"]/span[4]').text

        df_one = pd.DataFrame([{
            '来源作者': source_author, '期刊': journal, '年度(期)': issue,
            '标题': title, '作者': author, '机构': institute,
            '下载量': page_num, '页码': page_number, '页数': page_count, '大小': size
        }])

        file_exists = os.path.isfile(output_file)
        df_one.to_csv(output_file, mode='a', index=False, header=not file_exists, encoding='gbk')
    except:
        print("论文数据提取失败")
    finally:
        browser.close()
        browser.switch_to.window(browser.window_handles[2])  # 返回作者页


def process_author_page(browser, output_file):
    try:
        author_urls = browser.find_elements(By.XPATH, "//div[@class='wx-tit']/h3[@id='authorpart']/span/a")
        for author_url in author_urls:

            source_author = author_url.get_attribute('text')
            if source_author == '彭红枫1':
                continue
            url = author_url.get_attribute('href')
            browser.execute_script('window.open()')
            browser.switch_to.window(browser.window_handles[2])
            browser.get(url)
            time.sleep(3)

            last_height = browser.execute_script("return document.body.scrollHeight")
            while True:
                browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.5)
                new_height = browser.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

            while True:
                paper_list = browser.find_elements(By.XPATH, '//div[@id="KCMS-AUTHOR-JOURNAL-LITERATURES"]/ul[@class="ebBd"]/li')

                for p in paper_list:
                    try:
                        href3 = p.find_element(By.XPATH, "a").get_attribute('href')
                        browser.execute_script('window.open()')
                        browser.switch_to.window(browser.window_handles[3])
                        browser.get(href3)
                        extract_and_save_paper(browser, source_author, output_file)
                    except Exception as e:
                        print(f"单篇论文打开失败: {e}")

                try:
                    next_button = browser.find_element(By.XPATH, '//div[@id="KCMS-AUTHOR-JOURNAL-LITERATURES-page"]/a[@class="next"]')
                    browser.execute_script("arguments[0].click();", next_button)
                    time.sleep(2)
                except Exception as e:
                    print(f"点击下一页失败: {e}")
                    break

            browser.close()
            browser.switch_to.window(browser.window_handles[1])
            time.sleep(1)
    except:
        print("作者页面处理失败")


def main():
    input_file = './author/待爬虫数据0613.xlsx'
    output_file = './author/待爬虫数据0613_result.csv'
    df = pd.read_excel(input_file)

    browser = setup_browser()

    for idx, row in df.iterrows():
        journal = row['期刊']
        title = row['标题']
        print(f"正在处理：{idx}: {journal} - {title}")

        try:
            search_article(browser, journal, title)

            if open_article_detail(browser):
                process_author_page(browser, output_file)

                browser.close()
                browser.switch_to.window(browser.window_handles[0])
                browser.get("https://kns.cnki.net/kns8s/AdvSearch")
                time.sleep(1)
            else:
                print("检索失败，跳过该条")
        except Exception as e:
            print("出现异常，跳过当前条目：", e)
            browser.switch_to.window(browser.window_handles[0])
            browser.get("https://kns.cnki.net/kns8s/AdvSearch")
            time.sleep(1)

    browser.quit()


if __name__ == '__main__':
    main()
