from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import time
import datetime
import pandas as pd
from pandas import ExcelWriter
import urllib

pd.set_option('display.max_columns', 10)

# df = pd.DataFrame(columns=['Result_ID', 'Date', 'Time', 'Ping', 'Download_Speed', 'Upload_Speed', 'Server_Name', 'Server_Place', 'Result_URL'])
df = pd.read_excel('jio_speed_test.xls')
print df

browser = webdriver.Firefox()

while True:
    # Configure 200s timeout for loading of the speedtest.net page.
    try:
        browser.set_page_load_timeout(200)
        browser.get("http://www.speedtest.net/")

    except TimeoutException as ex:
        print("Exception has been thrown. " + str(ex))
        continue

    # Wait for the "GO" button to be rendered on the page.
    goClick = None
    while not goClick:
        try:
            goClick = browser.find_element_by_xpath(
                '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[1]/a/span[3]')
        except NoSuchElementException:
            # Wait for 2s for the button to finish being rendered.
            time.sleep(2)

    time.sleep(2)
    # Initiate the speed test by instrumenting a click on the "GO" button
    goClick.click()

    j = 0
    resultID = None
    resultTimeout = 250  # timeout in seconds to wait for a valid test result.
    # Wait for the test result to be produced.
    while not resultID:
        if j >= resultTimeout:
            break
        try:
            resultID = browser.find_element_by_xpath(
                '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[3]/div/div[1]/div[1]/div/div[2]/div[2]/a')

        except NoSuchElementException:
            time.sleep(1)
            j = j + 1
    if j >= resultTimeout:
        continue

    time.sleep(2)
    # Extract the resolved test result ID
    resultID = browser.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[3]/div/div[1]/div[1]/div/div[2]/div[2]/a')
    print "Result ID: {}".format(resultID.text)

    # Extract the download speed from the test results
    downspeed = browser.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[3]/div/div[1]/div[2]/div[2]/div/div[2]/span').text
    print "Download Speed: {}".format(downspeed)

    # Extract the upload speed from the test results
    upspeed = browser.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[3]/div/div[1]/div[2]/div[3]/div/div[2]/span').text
    print "Upload Speed: {}".format(upspeed)

    # Extract the ping latency from the test results
    pingg = browser.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[3]/div/div[1]/div[2]/div[1]/div/div[2]/span').text
    print "Ping: {}".format(pingg)

    # Extract the server name from the test results
    server_name = browser.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[4]/div/div[3]/div/div/div[2]/a').text
    print "Server Name: {}".format(server_name)

    # Extract the server location from the test results
    server_place = browser.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[3]/div/div[4]/div/div[3]/div/div/div[3]/span').text
    print "Server Location".format(server_place)

    # Format current timestamp
    ts = time.time()
    stamp_date = datetime.datetime.fromtimestamp(ts).strftime('%d-%m-%Y')
    stamp_time = datetime.datetime.fromtimestamp(ts).strftime('%H-%M')

    # Extract raw test results from sppedtest API and store it on disk
    urllib.urlretrieve("http://www.speedtest.net/result/" + resultID.text + ".png",
                       "jio_speedtest_ResultID-" + resultID.text + "_" + stamp_date + "_" + stamp_time + ".png")

    # Prepare result data for writing to .xls file.
    # df = pd.DataFrame(columns=['Result_ID', 'Date', 'Time', 'Ping', 'Download_Speed', 'Upload_Speed', 'Server_Name', 'Server_Place'])
    df = df.append({'Result_ID': resultID.text, 'Date': stamp_date, 'Time': stamp_time,
                    'Ping': pingg, 'Download_Speed': downspeed, 'Upload_Speed': upspeed,
                    'Server_Name': server_name, 'Server_Place': server_place,
                    'Result_URL': "http://www.speedtest.net/result/" + resultID.text + ".png"}, ignore_index=True)
    print df

    writer = ExcelWriter('jio_speed_test.xls')
    df.to_excel(writer, 'Sheet1')
    writer.save()

    time.sleep(120)

# browser.find_element_by_xpath("/html/body/div[3]/div[2]/div/div/div/div[3]/div[1]/div[1]/a/span[3]").click()
