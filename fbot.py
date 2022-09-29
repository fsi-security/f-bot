import os
import time
import boto3
import requests
import json
from bs4 import BeautifulSoup
from datetime import date
import urllib
import traceback
import win32gui, win32con, win32api 

TB_FSC_PRESS = "fscnotice"
TB_FSS_PRESS = "fssnotice"
TB_REPLY_CASE = "replycase"
TB_SANCTIONS = "sanctions"
TB_IMPROVE = "improves"
TB_SECURITY_BULLETINS = "security_bulletins"
TB_WHATS_NEW = "whats_new"
BETTER_BASE_URL = "https://better.fsc.go.kr"

def AWSWhatsNew(lang="ko_KR", page=0, size=10):
    
    applog("heartbeat - aws - AWSWhatsNew")        
    
    # ko_KR, en_US
    url = f"https://aws.amazon.com/api/dirs/items/search?item.directoryId=whats-new&sort_by=item.additionalFields.postDateTime&sort_order=desc&size={size}&item.locale={lang}&page={page}"    
    
    body = requests.request("POST", url).text   
    json_object = json.loads(body)
        
    for bulletin in json_object['items']:
        bulletin_item = bulletin['item']
        id = bulletin_item['id']
        bulletin_additional = bulletin_item['additionalFields']        
        headline = bulletin_additional['headline']
        headlineUrl = f"https://aws.amazon.com{bulletin_additional['headlineUrl']}"
        postDateTime = bulletin_additional['postDateTime']
        postBody = ""
        postSummary = ""
        try:
            postBody = bulletin_additional['postBody']
            postSummary = bulletin_additional['postSummary']
        except:
            pass
                
        month = postDateTime[0:7]
        storedObject = getArticle(TB_WHATS_NEW, "month", month, "id", id)
        if storedObject == None:
            newObject = {
                    "month": month,
                    "id": id,
                    "postDateTime": postDateTime,
                    "headline": headline,
                    "headlineUrl": headlineUrl,
                    "postBody": postBody,
                    "postSummary": postSummary,
                    "lang": lang
                }
            appendArticle(TB_WHATS_NEW, newObject)
            kakaomessage = f"[What's new][{postDateTime}]{headline}\n{headlineUrl}"
            send_message_to_me_kakao(kakaomessage)        
        

def SecurityBulletins(page=1):
    
    applog("heartbeat - aws - SecurityBulletins")
    
    size = 10
    url = f"https://aws.amazon.com/api/dirs/items/search?item.directoryId=security-bulletins&sort_by=item.additionalFields.bulletinId&sort_order=desc&size={size}&item.locale=en_US"
    
    body = requests.request("POST", url).text   
    json_object = json.loads(body)
        
    for bulletin in json_object['items']:
        bulletin_item = bulletin['item']
        bulletin_additional = bulletin_item['additionalFields']
        bulletin_id = bulletin_additional['bulletinId']
        bulletin_date = bulletin_additional['bulletinDateSort']
        subject_url = bulletin_additional['bulletinSubjectUrl']
        subject = bulletin_additional['bulletinSubject']
                
        year = bulletin_date[0:4]
        storedObject = getArticle(TB_SECURITY_BULLETINS, "year", year, "bulletin_id", bulletin_id)
        if storedObject == None:
            print (storedObject)
            newObject = {
                    "year": year,
                    "bulletin_id": bulletin_id,
                    "bulletin_date": bulletin_date,
                    "subject": subject,
                    "subject_url": subject_url
                }
            appendArticle(TB_SECURITY_BULLETINS, newObject)
            kakaomessage = f"[Security Bulletins][{bulletin_date}]{subject}\n{subject_url}"
            send_message_to_me_kakao(kakaomessage)
        
            

def fssNoticeList(page=1):

    applog("heartbeat - fss - noticeList")
    
    FSS_BASE_URL = "https://www.fss.or.kr"
    url = f"https://www.fss.or.kr/fss/bbs/B0000188/list.do?menuNo=200218&bbsId=&cl1Cd=&pageIndex={page}&sdate=&edate=&searchCnd=1&searchWrd="
    
    requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
    response = getHttpBody(url)
    
    soup = BeautifulSoup(response.text, 'html.parser')
    listBlock = soup.find("div", class_="bd-list")
    if listBlock == None:
        return
        
    items = listBlock.find_all("tr")
    
    for item in items:
        row = item.find_all("td")
        if len(row) < 4:
            continue
        day = row[3].text.strip()
        month = day[0:7]        

        count = item.find("td", class_="num").text.strip()
        titDiv = item.find("td", class_="title")
        
        storedObject = getArticle(TB_FSS_PRESS, "month", month, "count", int(count))        
        if storedObject == None:
            link = titDiv.find("a")['href'].strip()
            articleLink = f"{FSS_BASE_URL}{link}"

            attachment = fssNoticeDetail(count, day, articleLink)
            if attachment != None:
                object = {
                    "month": month,
                    "day": day,
                    "count": int(count),
                    "subject": attachment[0]['title'],
                    "articleLink": articleLink,
                    "division": attachment[0]['fields'][2]['value']
                }

                appendArticle(TB_FSS_PRESS, object)
                kakaomessage = f"[{object['day']}]{object['subject']} by {object['division']}\n{object['articleLink']}"
                send_message_to_me_kakao(kakaomessage)

def fssNoticeDetail(count, day, url):
        
    response = getHttpBody(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    viewDiv = soup.find("div", class_="bd-view")
    subjectDiv = viewDiv.find("h2", class_="subject")
    subject = subjectDiv.text.strip()
            
    divisionHeader = ""
    divisionValue = ""
    teamHeader = ""
    teamValue = ""    
    contactHeader = ""
    contactValue = ""
    content = ""
        
    if viewDiv != None:
            
        infoitems = viewDiv.find_all("dl")
        for item in infoitems:
            
            titBodies = item.find_all("dt")
            for body in titBodies:
                if body != None:
                    headerTitle = body.text.strip()
                    if headerTitle == "담당부서":
                        divisionHeader = headerTitle
                        divisionValue = body.findNext('dd').text.strip()
                        
                        maxlen = 100
                        contentDiv = body.parent.find_previous_sibling('div')
                        content = contentDiv.text.strip()
                        content = f"{content[:maxlen]}...[생략]..." if len(content) > maxlen else content                        
                    elif headerTitle == "담당팀":
                        teamHeader = headerTitle
                        teamValue = body.findNext('dd').text.strip()
                    elif headerTitle == "문의":
                        contactHeader = headerTitle
                        contactValue = body.findNext('dd').text.strip()        

    if "디지털" not in divisionValue \
        and "IT" not in divisionValue:
            # 중요하지 않은 부서는 본문 생략
            content = ""
    
    color = '#7CD197';
    # CRITICAL #ff0209, HIGH #ed7211, MEDIUM #ed7211, LOW #879596
    
    attachment = [{
        "fallback": "https://www.fss.go.kr",
        "pretext": "금융감독원 보도자료",
        "title": subject,
        "title_link": url,
        "text": content,
        "fields": [
            {"title": "등록일", "value": day, "short": "true"},
            {"title": contactHeader, "value": contactValue, "short": "true"},
            {"title": divisionHeader, "value": divisionValue, "short": "true"},
            {"title": teamHeader, "value": teamValue, "short": "true"}
        ],
        "mrkdwn_in": ["pretext"],
        "color": color
    }];

    return attachment

def auditList(cate="", page=1):

    if cate == "":
        applog("heartbeat - fsc - santionsList")
    else:
        applog("heartbeat - fsc - improveList")
    
    start_date = "2015-01-01"
    FSS_BASE_URL = "https://www.fss.or.kr"
    url = f"https://www.fss.or.kr/fss/job/openInfo{cate}/list.do?menuNo=200476&pageIndex={page}&sdate={start_date}&edate=&searchCnd=3&searchWrd="
    
    requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.text, 'html.parser')      

    listTable = soup.find("div", class_="bd-list")
    if listTable == None:
        return
        
    items = listTable.find_all("tr")
    if items == None:
        return
 
    for item in items:
        # 일련번호	제재대상기관	제재조치요구일	제재조치요구내용	관련부서	조회수
        tdBlock = item.find_all("td")
        if len(tdBlock) < 5:
            continue

        # 디지털 IT 핀테크
        idx = tdBlock[0].text.strip()
        targetComp = tdBlock[1].text.strip()
        reqDate = tdBlock[2].text.strip()
        year = reqDate[0:4]
        reqInfoBlock = tdBlock[3]
        reqInfoHref = ""
        reqDiv = tdBlock[4].text.strip()
        reqType = "검사결과제재" if cate == "" else "경영유의·개선사항"

        if "디지털" in reqDiv or "IT" in reqDiv or "핀테크" in reqDiv:
            if reqInfoBlock.find("a") != None and reqInfoBlock.find("a")['href'] != None:
                reqInfoHref = reqInfoBlock.find("a")['href'].strip()

                # ID추출
                detailLinkToken = urllib.parse.urlparse(reqInfoHref)
                detailLinkQuery = urllib.parse.parse_qs(detailLinkToken.query)
                examMgmtNo = detailLinkQuery['examMgmtNo'][0].strip()
                
                storedObject = getArticle(TB_SANCTIONS, "year", year, "exam_mgmt_no", int(examMgmtNo)) if cate == "" else getArticle(TB_IMPROVE, "year", year, "exam_mgmt_no", int(examMgmtNo))
                if storedObject == None:
                    reqInfoUrl = f"{FSS_BASE_URL}/fss/job/openInfo{cate}/{reqInfoHref[2:]}"
                    auditDetailInfo = auditDetail(reqInfoUrl)
                    auditDetailInfo["year"] = year
                    auditDetailInfo["exam_mgmt_no"] = int(examMgmtNo)
                    auditDetailInfo["req_date"] = reqDate
                    auditDetailInfo["req_type"] = reqType
                    auditDetailInfo["url"] = reqInfoUrl

                    message1 = f"[기관]{auditDetailInfo['sanctions_organization']}" if len(auditDetailInfo['sanctions_organization']) > 0 else ""
                    message2 = f"[임원]{auditDetailInfo['sanctions_director']}" if len(auditDetailInfo['sanctions_director']) > 0 else ""
                    message3 = f"[직원]{auditDetailInfo['sanctions_employee']}" if len(auditDetailInfo['sanctions_employee']) > 0 else ""
                    
                    message = message1
                    if len(message2) > 0:
                        if len(message) > 0:
                            message += "\n"
                        message += message2

                    if len(message3) > 0:
                        if len(message) > 0:
                            message += "\n"
                        message += message3
                    
                    attachment = [{
                        "fallback": "https://www.fss.or.kr",
                        "pretext": f"금융감독원 - {reqType}",
                        "title": auditDetailInfo['comp_Name'],
                        "title_link": reqInfoUrl,
                        "text": message,
                        "fields": [
                            {"title": "제재조치일", "value": auditDetailInfo['remediation_date'], "short": "true"},
                            {"title": "관련부서", "value": auditDetailInfo['relevant_dvision'], "short": "true"},
                            {"title": "첨부파일", "value": auditDetailInfo["attached_filename"], "short": "true"},
                        ],
                        "mrkdwn_in": ["pretext"],
                        "color": "#ed7211"
                    }];                    
                                                            
                    updatea_table = TB_SANCTIONS if cate == "" else TB_IMPROVE
                    appendArticle(updatea_table, auditDetailInfo)
                    
                    kakaomessage = f"[{attachment[0]['pretext']}]{attachment[0]['title']}\n{attachment[0]['title_link']}"
                    send_message_to_me_kakao(kakaomessage)


def auditDetail(url):
    
    soup = BeautifulSoup(requests.request("GET", url).text, 'html.parser')
    infoDiv = soup.select("div.bd-view")
    
    if infoDiv == None:
        return
    
    compTable = infoDiv[0]
    compTrBlockList = compTable.find_all("dl")

    response = {}

    for trBlock in compTrBlockList:
        thBlock = trBlock.find("dt")
        tdBlock = trBlock.find("dd")
        
        response["comp_Name"] = ""        
        response["remediation_date"] = ""
        response["relevant_dvision"] = ""
        response["attached_filename"] = ""
        
        if thBlock != None and thBlock.text.strip() == "금융기관명":
            compName = tdBlock.text.strip()
            response["comp_Name"] = compName
        elif thBlock != None and thBlock.text.strip() == "제재조치일":
            remediationDate = tdBlock.text.strip()
            response["remediation_date"] = remediationDate
        elif thBlock != None and thBlock.text.strip() == "관련부서":
            relevantDvision = tdBlock.text.strip()
            response["relevant_dvision"] = relevantDvision
        elif thBlock != None and "첨부파일" in thBlock.text.strip():
            attachedFilename = tdBlock.text.strip()
            response["attached_filename"] = attachedFilename
   
    if len(infoDiv) > 1:
        infoTable = infoDiv[1].find_all("dd")
        
        if infoTable != None and len(infoTable) >= 3:
            response["sanctions_organization"] = infoTable[0].text.strip()
            response["sanctions_director"] = infoTable[1].text.strip()
            response["sanctions_employee"] = infoTable[2].text.strip()

    return response


def replyCaseList(page=0):

    applog("heartbeat - fsc - replyCaseList")
    
    start_date = "2015-01-01"
    category = 4 #(4)전자금융
    size = 10
    url = f"https://better.fsc.go.kr/fsc_new/replyCase/selectReplyCasePastReplyList.do?draw=1&columns[0][data]=rownumber&columns[0][name]=&columns[0][searchable]=true&columns[0][orderable]=false&columns[0][search][value]=&columns[0][search][regex]=false&columns[1][data]=gubun&columns[1][name]=&columns[1][searchable]=true&columns[1][orderable]=false&columns[1][search][value]=&columns[1][search][regex]=false&columns[2][data]=category&columns[2][name]=&columns[2][searchable]=true&columns[2][orderable]=false&columns[2][search][value]=&columns[2][search][regex]=false&columns[3][data]=title&columns[3][name]=&columns[3][searchable]=true&columns[3][orderable]=false&columns[3][search][value]=&columns[3][search][regex]=false&columns[4][data]=number&columns[4][name]=&columns[4][searchable]=true&columns[4][orderable]=false&columns[4][search][value]=&columns[4][search][regex]=false&columns[5][data]=regDate&columns[5][name]=&columns[5][searchable]=true&columns[5][orderable]=false&columns[5][search][value]=&columns[5][search][regex]=false&order[0][column]=0&order[0][dir]=asc&start={page * size}&length=10&search[value]=&search[regex]=&searchKeyword=&searchCondition=&searchReplyRegDateStart={start_date}&searchReplyRegDateEnd=&searchType=&searchCategory={category}&searchLawType="
    
    body = requests.request("POST", url).text   
    json_object = json.loads(body)
        
    for object in json_object['data']:
        idx = object['idx']
        number = object['number'] #일련번호
        regDate = object['regDate'] #등록일
        year = regDate[0:4]
        gubun = "1" if object['gubun'] == "법령해석" else "2"
        
        requestLawreqUrl = f"{BETTER_BASE_URL}/fsc_new/replyCase/LawreqDetail.do?muNo=171&stNo=11&lawreqIdx={idx}&actCd=R"
        requestOpinionDetailUrl = f"{BETTER_BASE_URL}/fsc_new/replyCase/OpinionDetail.do?muNo=171&stNo=11&opinionIdx={idx}&actCd=R"
        requestUrl = requestLawreqUrl if gubun == "1" else requestOpinionDetailUrl
        
        storedObject = getArticle(TB_REPLY_CASE, "year", year, "idx", int(idx))
        if storedObject == None:
            replyCaseDetailInfo = replyCaseDetail(requestUrl)
            replyCaseDetailInfo['year'] = year
            replyCaseDetailInfo['idx'] = idx # int타입
            replyCaseDetailInfo['number'] = number
            replyCaseDetailInfo['regDate'] = regDate
            replyCaseDetailInfo['gubun'] = gubun
            replyCaseDetailInfo['requestUrl'] = requestUrl
            
            if replyCaseDetailInfo['result'] == "완료":
                message1 = f"[질의요지]{replyCaseDetailInfo['keynote']}" if len(replyCaseDetailInfo['keynote']) > 0 else ""
                message2 = f"[회답]{replyCaseDetailInfo['reply']}" if len(replyCaseDetailInfo['reply']) > 0 else ""
                message3 = f"[이유]{replyCaseDetailInfo['reason']}" if len(replyCaseDetailInfo['reason']) > 0 else ""
                message = message1
                if len(message2) > 0:
                    if len(message) > 0:
                        message += "\n"
                    message += message2
                if len(message3) > 0:
                    if len(message) > 0:
                        message += "\n"
                    message += message3
                
                pretext = "금융위원회 법령해석" if gubun == "1" else "금융위원회 비조치의견서" 
                attachment = [{
                    "fallback": "https://better.fsc.go.kr",
                    "pretext": pretext,
                    "title": replyCaseDetailInfo['subject'],
                    "title_link": requestUrl,
                    "text": message,
                    "fields": [
                        {"title": "등록일", "value": replyCaseDetailInfo['regDate'], "short": "true"},
                        {"title": "회신일", "value": replyCaseDetailInfo['responseDate'], "short": "true"},
                        {"title": "등록자", "value": replyCaseDetailInfo['register'], "short": "true"},
                        {"title": "첨부파일", "value": replyCaseDetailInfo['attachedFileName'], "short": "true"}
                    ],
                    "mrkdwn_in": ["pretext"],
                    "color": "#ff0209"
                }];
                appendArticle(TB_REPLY_CASE, replyCaseDetailInfo)
                kakaomessage = f"[{attachment[0]['pretext']}]{attachment[0]['title']}\n{attachment[0]['title_link']}"
                send_message_to_me_kakao(kakaomessage)

def replyCaseDetail(url):
    
    soup = BeautifulSoup(requests.request("POST", url).text, 'html.parser')
    
    subjectTable = soup.find("table", class_="tbl-view")
    subject = subjectTable.find("td", class_="subject").text.strip()
    
    tableDiv = soup.find("table", class_="tbl-write")
    trDivList = tableDiv.find_all("tr")   

    response = {}
    response["subject"] = subject
    
    for trDiv in trDivList:
        thDiv = trDiv.find("th")
        tdDiv = trDiv.find("td")
        
        if thDiv != None and thDiv.text.strip() == "첨부파일":
            attachedFileName = tdDiv.text.strip()
            if tdDiv.find("a") == None or tdDiv.find("a")['href'] == None:
                response["attachedFileName"] = "파일없음"
                continue
           
            attachedHref = tdDiv.find("a")['href'].strip()
            attachedLink = f"{BETTER_BASE_URL}{attachedHref}" #link에 앞에 / 붙어서 옴
            
            # 첨부파일명한글 때문에 urlencoding
            attachedLinkToken = urllib.parse.urlparse(attachedLink)
            attachedLinkQuery = urllib.parse.parse_qs(attachedLinkToken.query)
            encodedAttachedLinkQuery = urllib.parse.urlencode(attachedLinkQuery, doseq=True)
            attachedFileLink = f"{attachedLinkToken.scheme}://{attachedLinkToken.netloc}{attachedLinkToken.path}?{encodedAttachedLinkQuery}"           
            response["attachedFileName"] = attachedFileName
            response["attachedFileLink"] = attachedFileLink
        elif thDiv != None and thDiv.text.strip() == "등록자":
            register = tdDiv.text.strip()
            response["register"] = register
        elif thDiv != None and thDiv.text.strip() == "회신일":
            responseDate = tdDiv.text.strip()
            response["responseDate"] = responseDate
        elif thDiv != None and thDiv.text.strip() == "질의요지":
            keynote = tdDiv.text.strip()
            response["keynote"] = keynote if len(keynote) < 100 else f"{keynote[:100]}..이하생략.."
        elif thDiv != None and thDiv.text.strip() == "회답":
            reply = tdDiv.text.strip()
            response["reply"] = reply if len(reply) < 100 else f"{reply[:100]}..이하생략.."
        elif thDiv != None and thDiv.text.strip() == "이유":
            reason = tdDiv.text.strip()
            response["reason"] = reason if len(reason) < 100 else f"{reason[:100]}..이하생략.."
        elif thDiv != None and thDiv.text.strip() == "처리구분":
            result = tdDiv.text.strip()
            response["result"] = tdDiv.text.strip()
        
    return response    

def fscNoticeList(page=1):

    FSC_BASE_URL = "https://www.fsc.go.kr"
    url = f"https://www.fsc.go.kr/no010101?curPage={page}"
    
    requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
    response = requests.post(url, verify=False)
    
    soup = BeautifulSoup(response.text, 'html.parser')

    listBlock = soup.find("div", class_="board-list-wrap")
    items = listBlock.find_all("div", class_="inner")
    
    applog("heartbeat - fsc - noticeList")
    
    for item in items:
        countDiv = item.find("div", class_="count")
        count = countDiv.text.strip()        
        
        dayDiv = item.find("div", class_="day")
        day = dayDiv.text.strip()
        month = day[0:7]
        
        storedObject = getArticle(TB_FSC_PRESS, "month", month, "count", int(count))        
        if storedObject == None:
            link = item.find("a")['href'].strip()
            articleLink = "%s%s" %(FSC_BASE_URL, link) 
            
            attachment = fscNoticeDetail(count, day, articleLink)
            if attachment != None:
                object = {
                    "month": month,
                    "day": day,
                    "count": int(count),
                    "subject": attachment[0]['title'],
                    "articleLink": articleLink,
                    "division": attachment[0]['fields'][2]['value']
                }

                appendArticle(TB_FSC_PRESS, object)
                kakaomessage = f"[{object['day']}]{object['subject']} by {object['division']}\n{object['articleLink']}"               
                send_message_to_me_kakao(kakaomessage)

def fscNoticeDetail(count, day, url):
        
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.text, 'html.parser')
    viewDiv = soup.find("div", class_="board-view-wrap")
   
    subjectDiv = viewDiv.find("div", class_="subject")
    subject = subjectDiv.text.strip()    
    
    infoDiv = viewDiv.find("div", class_="info")
    if infoDiv != None:
        infoitems = infoDiv.find_all("span")

        divisionHeader = infoitems[0].find("strong").text.strip() if len(infoitems) > 0 else ""
        divisionValue = infoitems[0].text.strip().replace(divisionHeader, "") if len(infoitems) > 0  else ""

        managerHeader = infoitems[1].find("strong").text.strip() if len(infoitems) > 1 else ""
        managerValue = infoitems[1].text.strip().replace(managerHeader, "") if len(infoitems) > 1  else ""    

        contactHeader = infoitems[2].find("strong").text.strip() if len(infoitems) > 2 else ""
        contactValue = infoitems[2].text.strip().replace(contactHeader, "") if len(infoitems) > 2  else ""

        maxlen = 100
        contentDiv = viewDiv.find("div", class_="cont")
        content = contentDiv.text.strip()
        content = f"{content[:maxlen]}...[생략]..." if len(content) > maxlen else content
    else: # 없음
        divisionHeader = ""
        divisionValue = ""
        managerHeader = ""
        managerValue = ""    
        contactHeader = ""
        contactValue = ""
        content = ""

    if "전자금융" not in divisionValue \
        and "금융혁신" not in divisionValue \
        and "금융데이터" not in divisionValue \
        and "FIU" not in divisionValue \
        and "샌드박스" not in divisionValue:
            # 중요하지 않은 부서는 본문 생략
            content = ""
    
    color = '#7CD197';
    # CRITICAL #ff0209, HIGH #ed7211, MEDIUM #ed7211, LOW #879596
    
    attachment = [{
        "fallback": "https://www.fsc.go.kr",
        "pretext": "금융위원회 보도자료",
        "title": subject,
        "title_link": url,
        "text": content,
        "fields": [
            {"title": "등록일", "value": day, "short": "true"},
            {"title": contactHeader, "value": contactValue, "short": "true"},
            {"title": divisionHeader, "value": divisionValue, "short": "true"},
            {"title": managerHeader, "value": managerValue, "short": "true"}
        ],
        "mrkdwn_in": ["pretext"],
        "color": color
    }];

    return attachment

def getHttpBody(url):
    
    requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
    response = requests.post(url, verify=False)
    
    return response
    
def getArticle(tableName, month_h, month_v, count_h, count_v):

    dynamodb = boto3.resource('dynamodb', region_name='ap-northeast-2', endpoint_url="http://dynamodb.ap-northeast-2.amazonaws.com")
    table = dynamodb.Table(tableName) 

    try:
        response = table.get_item(Key={month_h: month_v, count_h: count_v})
    except Exception as e:
        print(traceback.format_exc())
    else:
        if "Item" in response:
            return response['Item']
        else:
            return None    

def appendArticle(tableName, object):
    
    dynamodb = boto3.resource('dynamodb', region_name='ap-northeast-2', endpoint_url="http://dynamodb.ap-northeast-2.amazonaws.com")
    table = dynamodb.Table(tableName)
    response = table.put_item(Item=object)
    
    return response


def send_message_to_me_kakao(attachment):

    kakao_room_title = "lee"
    
    kakao = win32gui.FindWindow(None, kakao_room_title) 
    chat = win32gui.FindWindowEx(kakao, None , "RICHEDIT50W" , None) # 채팅창안 메세지 입력창 

    cText = attachment 
    applog(cText)
    win32api.SendMessage(chat, win32con.WM_SETTEXT, 0, cText) # 채팅창 입력 
    win32api.PostMessage(chat, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0) 
    win32api.PostMessage(chat, win32con.WM_KEYUP, win32con.VK_RETURN, 0) # 엔터키
    
    # delay 5초
    time.sleep(5)


def applog(msg):
    print(f"[{getCurrentTime()}]{msg}")


def getCurrentTime():    
    now = time.localtime()
    return "%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
    
if __name__ == "__main__":
    print ('start...')
    while True:
        try:            
            fscNoticeList() #금융위보도자료
            fssNoticeList() #금감원보도자료
            replyCaseList() #비조치의견,법령해석
            auditList("") #제재관련 공시(blank)
            auditList("Impr") #경영유의사항 등 공시(/impr)
            SecurityBulletins() #AWS 보안
            # AWSWhatsNew("ko_KR", 0, 10) #AWS What's New
        except Exception as e:
            print(traceback.format_exc())
        
        time.sleep(1800)
