from docx import Document

list = [
    "デイリーポータルＺライター賞",
    "爲房新太朗様",
    "あなたの深く考えられた記事と、緻密で愛情ある文章に感銘を受け、デイリーポータルＺライター賞を授与いたします。あなたの寄稿は常に読者の注目を集め、彼らに新たな視点を提供しています。更なるご成功を心よりお祈りしております。",
    "授与日：2021年7月10日",
    "授与組織：日本メディア協会"
]

doc = Document("template.docx")
print(doc.paragraphs[2].text)
print(doc.paragraphs[2].runs[0].text)
doc.paragraphs[0].runs[0].text = "置換したテキストだよ"
doc.save("sample.docx")
# doc.add_paragraph(textwrap.dedent('''デイリーポータルＺライター賞

# 爲房新太朗様

# あなたの深く考えられた記事と、緻密で愛情ある文章に感銘を受け、デイリーポータルＺライター賞を授与いたします。あなたの寄稿は常に読者の注目を集め、彼らに新たな視点を提供しています。更なるご成功を心よりお祈りしております。

# 授与日：2021年7月10日
# 授与組織：日本メディア協会'''))
#doc.add_paragraph(list[0])
#doc.save("sample.docx")