import openai

file = open(".apikey")
openai.api_key = file.readline()
file.close()

order = "爲房新太朗は、本業のWebエンジニアの業務もある中でライターとしてデイリーポータルＺに記事を寄稿しています。そのことに対して架空の賞を設定し、さらに賞状を贈呈するとします。「賞の名前」「その賞状に記載する文章」「授与する組織名」を考えて、json形式で回答してください。jsonのkeyは、「賞の名前」はaward、「その賞状に記載する文章」はtext、「授与する組織名」はorgとしてください。賞状に記載する文章は勝手に考えてよく、200文字程度としてください。"

#completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=[{"role": "user", "content": order}])
completion = openai.ChatCompletion.create(model="gpt-4", messages=[{"role": "user", "content": order}],   temperature=1, max_tokens=256, top_p=1, frequency_penalty=0, presence_penalty=0)
print(completion.choices[0].message.content)