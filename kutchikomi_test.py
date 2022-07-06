from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import gspread
import time
import os
import csv

#Googleスプレッドシートの操作

wb = gspread.service_account(filename="./syukatstu-4c493bb05893.json")
ws = wb.open_by_key("1GTnbDXHyMqe9ISry2Eifw9uwhArKM7xe3Ul87xTr1LA")
sh = ws.get_worksheet(0)

#A列の企業名だけを取得する
kigyouname_ls = sh.col_values(1)
kigyouname_ls = kigyouname_ls[1:]

#ブラウザの起動
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.99 Safari/537.36')
browser = webdriver.Chrome(ChromeDriverManager().install(),options = options)

i = 2
#一つずつ企業の口コミを取得していくよ！
for KEYWORD in kigyouname_ls:
    try:   
        browser.get('https://careerconnection.jp/')
        time.sleep(0.5)
        #検索ボックスへの入力
        browser.find_element_by_css_selector('#SUGGEST_INPUT').send_keys(KEYWORD)
        #検索ボタンをクリック
        browser.find_element_by_css_selector('#corpButton').click()
        #検索結果の一番上のものを表示させる。
        companyname = browser.find_element_by_css_selector('#company_result > ul > li:nth-child(1) > h2 > a').text
        
        if KEYWORD.lower() in companyname.lower():
            browser.get(browser.find_element_by_css_selector('#company_result > ul > li:nth-child(1) > h2 > a').get_attribute('href'))
            time.sleep(0.2)
            career_connect_ls = []
            #総合評価の取得
            Assessment = browser.find_element_by_css_selector('#main > article > div > section.pc-review-corpseq-overview > div > div:nth-child(2) > dl').text.replace('\n',' ')
            career_connect_ls.append(Assessment)
            #各評価の取得
            Each_evaluation_elem = browser.find_element_by_css_selector('#canvas_detail')
            working_time = Each_evaluation_elem.get_attribute('data-chartl_1') + ' ' + Each_evaluation_elem.get_attribute('data-chart1')
            Rewarding = Each_evaluation_elem.get_attribute('data-chartl_2') + ' ' + Each_evaluation_elem.get_attribute('data-chart2')
            Stress_level = Each_evaluation_elem.get_attribute('data-chartl_3') + ' ' + Each_evaluation_elem.get_attribute('data-chart3')
            holiday = Each_evaluation_elem.get_attribute('data-chartl_4') + ' ' + Each_evaluation_elem.get_attribute('data-chart4')
            Salary = Each_evaluation_elem.get_attribute('data-chartl_5') + ' ' + Each_evaluation_elem.get_attribute('data-chart5')
            Whiteness = Each_evaluation_elem.get_attribute('data-chartl_5') + ' ' + Each_evaluation_elem.get_attribute('data-chart5')
            career_connect_ls.append(working_time+'\n'+working_time+'\n'+Rewarding+'\n'+Stress_level+'\n'+holiday+'\n'+Salary+'\n'+Whiteness)

            #平均年収
            income_av = browser.find_element_by_css_selector('#main > article > div > section.pc-review-corpseq-overview > div > div:nth-child(2) > div.overview-area__income > dl.overview-area__income-list1').text.replace('\n',' ')
            career_connect_ls.append(income_av)

            sh.update('B'+str(i)+':D'+str(i),[career_connect_ls])
        else:
            sh.update('B'+str(i),'該当結果無し。')
            print('検索がヒットしませんでした。')
            
    except:
        print('キャリコネからはうまく取得できませんでした。')
        sh.update('B'+str(i),'該当結果無し。')

    #転職会議から同様に取得する
    try:
        browser.get('https://jobtalk.jp/')
        time.sleep(0.2)
        browser.find_element_by_css_selector('#__next > div.css-o3m7gb > div > div > div > div > form > input.css-10oqfui').send_keys(KEYWORD)
        browser.find_element_by_css_selector('#__next > div.css-o3m7gb > div > div > div > div > form > button').click()
        time.sleep(0.2)

        #検索結果の一番上の詳細を開く。
        search_1_elem = browser.find_element_by_css_selector('#__next > div > div.PcTemplate_container__2KiM8 > div.PcTemplate_mainColumn__19miA > div.CompanySearchResultList_wrapper__1xLBf > div:nth-child(1) > div.CompanySearchResultHeader_companyHeader__vH4MF > div > h2 > span > a')
        if KEYWORD in mojimoji.zen_to_han(search_1_elem.text,kana=False):
            browser.get(search_1_elem.get_attribute('href'))
            syusyoku_mtg_ls = []
            #総合評価　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　
            syusyoku_mtg_ls.append(browser.find_element_by_css_selector('#__next > div > div.PcTemplate_container__Yv6GG.PcTemplate_containerMain__3q6_r > div.PcTemplate_containerMainLeft__18SKy > div > div.PcTemplate_sectionContentMargin__3sNe1 > div > div > div > div > div.CompanyScoreSection_rating__2whG1').text.replace('\n',' '))
            #平均年収
            syusyoku_mtg_ls.append('平均'+browser.find_element_by_css_selector('#__next > div > div.PcTemplate_container__Yv6GG.PcTemplate_containerMain__3q6_r > div.PcTemplate_containerMainLeft__18SKy > div > div.PcTemplate_salaryScoreSectionMargin__26m0K > div > div.SalaryStatPcSection_title__3FbeM > div').text.replace('\n',' '))
            sh.update('E'+str(i)+':F'+str(i),[syusyoku_mtg_ls])
        else:
            sh.update('E'+str(i),'該当結果無し。')
    except:
        print('就職会議からはうまく取得できませんでした。')
        sh.update('E'+str(i),'該当結果無し。')

    try:
        #オープンワークスからスクレイピング
        browser.get('https://www.vorkers.com/')
        time.sleep(0.2)
        #検索ボックスへ入力
        browser.find_element_by_css_selector('#jsTopSearchCompany > form > div > div > input').send_keys(KEYWORD)
        #検索ボタンをクリック
        browser.find_element_by_css_selector('#jsTopSearchCompany > form > div > div > button').click()
        time.sleep(0.2)
        #一番上の検索結果を取得する
        if KEYWORD.lower() in browser.find_element_by_css_selector('#contentsBody > section > ul > li:nth-child(1) > div.searchCompanyName > div:nth-child(1) > h3').text.lower():
            browser.get(browser.find_element_by_css_selector('#contentsBody > section > ul > li:nth-child(1) > div.searchCompanyName > div:nth-child(1) > h3 > a').get_attribute('href'))

            #必要情報を取得
            openworks_ls = []
            total_score = '総合得点' + browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_right.averageScore_right-companyTop > div.mt-5 > p.totalEvaluation_item.fs-17 > span').text
            openworks_ls.append(total_score)
            Treatment_aspect = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-satisfy').text.replace('\n',' ')
            spirit_score = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-spirit').text.replace('\n',' ')
            airy_score = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-airy').text.replace('\n',' ')
            team_score = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-team').text.replace('\n',' ')
            junior_score = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-junior').text.replace('\n',' ')
            senior_score = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-senior').text.replace('\n',' ')
            law_score = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-law').text.replace('\n',' ')
            assess = browser.find_element_by_css_selector('#mainColumn > article:nth-child(1) > div.averageScore > div.averageScore_chart.averageScore_chart-companyTop > ul > li.scoreList_item-assess').text.replace('\n',' ')
            openworks_ls.append(Treatment_aspect+'\n'+spirit_score+'\n'+airy_score+'\n'+team_score+'\n'+junior_score+'\n'+senior_score+'\n'+law_score)
            income_av = '平均年収 ' + browser.find_element_by_css_selector('#mainColumn > article:nth-child(3) > div.borderGray.pl-20.pr-20.pt-25.pb-25 > table > tbody > tr.d-b.mt-n5 > td').text
            openworks_ls.append(income_av)

            sh.update('G'+str(i)+':I'+str(i),[openworks_ls])
        else:
            sh.update('G'+str(i),'該当結果無し。')
    except:
        print('オープンワークスからはうまく取得できませんでした。')
        sh.update('G'+str(i),'該当結果無し。')

    i += 1


with open('test.csv','w') as f:
    csv.writer(f).writerows(sh.get_all_values())

print('すべての処理が完了しました。')