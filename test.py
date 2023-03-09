import openai

file = open(".apikey")
openai.api_key = file.readline()
file.close()

order = "爲房新太朗は、本業のWebエンジニアの業務もある中でライターとしてデイリーポータルＺに記事を寄稿しています。そのことに対して架空の賞を設定し、さらに賞状を贈呈するとして、その賞状に記載する文章を考えてください。受賞した理由は勝手に考えて構いません。また、賞状の文章は1行目に賞の名前、2行目に受賞者名、3行目以降に残りの文章を記載するようにしてください。受賞者名の敬称は「様」としてください。末尾には受賞日として今日の日付と、授与する組織名を捏造して記載してください。また「賞の名前」「受賞者名」「受賞日」「授与する組織名」といったラベルは生成する文章には含めず、またプレースホルダとなる要素は必要ありません。例示した文章をそのまま採用しコピー＆ペーストして利用できるようにしてください。"

completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=[{"role": "user", "content": order}])
print(completion.choices[0].message.content)