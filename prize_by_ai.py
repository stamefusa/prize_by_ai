from docx import Document
import openai
import datetime

def main():
    prompt = "爲房新太朗は、おいしいケーキを作ることに長けており、家族をそのケーキで幸せにしています。そのことに対して架空の賞を設定し、さらに賞状を贈呈するとして、賞の名前・その賞状に記載する文章・授与する組織名を考えて、その順番にタブ区切りで回答してください。賞状に記載する文章は勝手に考えてよく、200文字以上400文字以内としてください。"
    prize = getPrizeTexts(prompt)
    #prize = ['美味しいケーキ賞', '太朗くんのおいしいケーキで家族を幸せにする功績をたたえます。これからも、おいしいケーキで人々を幸せにしてください。', '架空の「お菓子協会」']
    print(prize)

    if len(prize) == 3:
        print("成功")
        now = datetime.datetime.now()
        disp_now = str(now.year) + "年" + str(now.month) + "月" + str(now.day) + "日"
        makeDocx(prize[0], "爲房 新太朗", prize[1], disp_now, prize[2])
    else:
        print("失敗")

def getPrizeTexts(input):
    file = open(".apikey")
    openai.api_key = file.readline()
    file.close()
    
    completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=[{"role": "user", "content": input}])
    tmp = completion.choices[0].message.content
    #print(tmp)
    #print(tmp.strip().split("\t"))
    return tmp.strip().split("\t")

def makeDocx(title, to_name, body, date, from_name):
    doc = Document("template.docx")
    doc.paragraphs[0].runs[0].text = title
    doc.paragraphs[2].runs[0].text = to_name
    doc.paragraphs[4].runs[0].text = body
    doc.paragraphs[6].runs[0].text = date
    doc.paragraphs[7].runs[0].text = from_name
    doc.save("sample.docx")

if __name__ == "__main__":
    main()