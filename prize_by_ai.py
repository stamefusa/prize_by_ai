from docx import Document
import openai
import datetime
import subprocess
import os
import re

def main():
    name, text = getInputTexts()
    print(name + "\n" + text)
    prompt = name + "は、" + text + "。そのことに対して架空の賞を設定し、さらに賞状を贈呈するとして、賞の名前・その賞状に記載する文章・授与する組織名を考えて、その順番にタブ区切りで回答してください。賞状に記載する文章は勝手に考えてよく、300文字程度としてください。"
    prize = getPrizeTexts(prompt)
    print(prize)

    if len(prize) == 3:
        print("成功")
        now = datetime.datetime.now()
        disp_now = "令和" + str(now.year-2018) + "年" + str(now.month) + "月" + str(now.day) + "日"
        title, body, org = sanitise(prize)
        makeDocx(title, name, body, disp_now, org)
        printPrize()
    else:
        print("失敗")

def getInputTexts():
    file = open("input.txt")
    name, text = file.read().splitlines()
    return name, text.rstrip("。")

def getPrizeTexts(input):
    file = open(".apikey")
    openai.api_key = file.readline()
    file.close()
    
    completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=[{"role": "user", "content": input}])
    tmp = completion.choices[0].message.content
    return tmp.strip().split("\t")

def sanitise(src):
    # 各要素で改行、「」は削除。また、「賞の名前：」といったラベルをつけることがあるのでそれも削除。
    # 賞の文章のみ、句点で改行する。

    # 賞の名前
    title = re.sub(r'\n|.*：|「|」', '', src[0])

    # 賞の文章
    body = re.sub(r'\n|.*：|「|」', '', src[1]).replace('。', '。\n')

    # 授与する組織名
    org = re.sub(r'\n|.*：|「|」', '', src[2])

    return title, body, org

def makeDocx(title, to_name, body, date, from_name):
    # TODO 賞の文章の長さによって受賞者名の位置を調整
    doc = Document("template.docx")
    doc.paragraphs[0].runs[0].text = title
    doc.paragraphs[2].runs[0].text = to_name
    doc.paragraphs[4].runs[0].text = body
    doc.paragraphs[6].runs[0].text = date
    doc.paragraphs[7].runs[0].text = from_name
    doc.save("sample.docx")

def printPrize():
    path = os.getcwd() + "/sample.docx"
    subprocess.run("docx2pdf " + path, shell=True)
    print_path = os.getcwd() + "/sample.pdf"
    #subprocess.run("lpr -P EPSON_EP_707A_Series " + print_path, shell=True)

if __name__ == "__main__":
    main()