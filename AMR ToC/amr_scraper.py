from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import csv
from time import sleep
import random
import os
import json

def setup_driver():
    edge_options = Options()
    
    # 添加 headless 模式的选项
    edge_options.add_argument('--headless')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--disable-extensions')
    edge_options.add_argument('--disable-software-rasterizer')
    edge_options.add_argument('--ignore-certificate-errors')
    edge_options.add_argument('--window-size=1920,1080')
    edge_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
    
    # 添加更多选项来抑制警告和错误信息
    edge_options.add_argument('--log-level=3')  # 只显示致命错误
    edge_options.add_argument('--silent')
    edge_options.add_argument('--disable-logging')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    edge_options.add_experimental_option('useAutomationExtension', False)
    
    try:
        service = Service(r"D:\msedgedriver.exe")
        # 添加service_args来抑制WebDriver日志
        service.creation_flags = 0x08000000  # CREATE_NO_WINDOW
        driver = webdriver.Edge(service=service, options=edge_options)
        return driver
    except Exception as e:
        print(f"启动Edge浏览器时出错: {str(e)}")
        print("请确保：")
        print("1. Edge浏览器已正确安装")
        print("2. msedgedriver.exe 版本与Edge浏览器版本匹配")
        print("3. msedgedriver.exe 路径正确")
        raise e

def write_to_csv(articles, file_path, is_new=False):
    mode = 'w' if is_new else 'a'
    with open(file_path, mode, newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['Year', 'Volume', 'Issue', 'Title', 'Authors', 'DOI', 'Link'])
        if is_new:  # 只在新文件时写入表头
            writer.writeheader()
        writer.writerows(articles)

def get_year_from_volume(volume):
    # AMR 第1卷是1976年，每卷加一年
    return 1976 + (volume - 1)

def scrape_amr_page(volume, issue):
    driver = setup_driver()
    
    try:
        url = f"https://journals.aom.org/toc/amr/{volume}/{issue}"
        print(f"\n正在访问页面: {url}")
        driver.get(url)
        
        # 等待页面加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "article-meta"))
        )
        
        # 给页面一点额外时间完全加载
        sleep(random.uniform(1, 2))
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        articles = []
        
        article_elements = soup.find_all('div', class_='article-meta')
        print(f"找到 {len(article_elements)} 个文章元素")
        
        for article in article_elements:
            try:
                title_link = article.find('h5', class_='issue-item__title').find('a')
                title = title_link.text.strip()
                doi = title_link['href'].replace('/doi/', '')
                
                # 根据卷号计算年份
                year = get_year_from_volume(volume)
                
                authors_list = article.find('ul', class_='rlist--inline loa')
                if authors_list:
                    authors = [author.find('span').text.strip() for author in authors_list.find_all('li')]
                    authors = '; '.join(authors)
                else:
                    authors = ''
                
                link = f"https://journals.aom.org{title_link['href']}"
                
                articles.append({
                    'Year': year,
                    'Volume': volume,
                    'Issue': issue,
                    'Title': title,
                    'Authors': authors,
                    'DOI': doi,
                    'Link': link
                })
                
            except Exception as e:
                print(f"解析文章时出错: {str(e)}")
                continue
        
        return articles
        
    except Exception as e:
        print(f"爬取 Volume {volume}, Issue {issue} 时出错: {str(e)}")
        return []
    
    finally:
        driver.quit()

def load_checkpoint():
    checkpoint_file = 'AMR ToC/checkpoint.json'
    if os.path.exists(checkpoint_file):
        with open(checkpoint_file, 'r') as f:
            return json.load(f)
    return {'last_volume': 1, 'last_issue': 0}  # 默认从第1卷开始

def save_checkpoint(volume, issue):
    checkpoint_file = 'AMR ToC/checkpoint.json'
    with open(checkpoint_file, 'w') as f:
        json.dump({'last_volume': volume, 'last_issue': issue}, f)

def main():
    output_file = 'AMR ToC/amr_articles.csv'
    total_articles = 0
    
    # 检查是否需要创建新文件
    is_new_file = not os.path.exists(output_file)
    
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # 加载断点信息
    checkpoint = load_checkpoint()
    start_volume = checkpoint['last_volume']
    start_issue = checkpoint['last_issue'] + 1
    if start_issue > 4:  # 如果上次完成了某卷的最后一期
        start_volume += 1
        start_issue = 1
    
    try:
        for volume in range(start_volume, 50):  # 修改为爬取到第49卷 (range的结束值要比目标大1)
            for issue in range(start_issue, 5):
                print(f"\n正在爬取第 {volume} 卷第 {issue} 期...")
                articles = scrape_amr_page(volume, issue)
                
                if articles:
                    write_to_csv(articles, output_file, is_new=(is_new_file and total_articles == 0))
                    total_articles += len(articles)
                    print(f"本期成功爬取 {len(articles)} 篇文章")
                    print(f"当前总计已爬取 {total_articles} 篇文章")
                    
                    latest_article = articles[-1]
                    print("\n最新爬取文章信息：")
                    print(f"年份：{latest_article['Year']}")
                    print(f"标题：{latest_article['Title']}")
                    print(f"作者：{latest_article['Authors']}")
                    print(f"DOI：{latest_article['DOI']}")
                
                # 保存断点
                save_checkpoint(volume, issue)
                
                # 增加随机延迟到1-3秒
                sleep_time = random.uniform(1, 2)
                print(f"等待 {sleep_time:.1f} 秒...")
                sleep(sleep_time)
            
            # 重置下一卷的起始期数为1
            start_issue = 1
        
        print(f"\n爬取完成！共爬取 {total_articles} 篇文章")
        print(f"数据已保存到 {output_file}")
        
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        # 保存当前进度
        save_checkpoint(volume, issue)
        print("已保存断点信息，下次运行从此处继续")

if __name__ == "__main__":
    main() 