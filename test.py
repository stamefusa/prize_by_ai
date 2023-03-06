import openai

file = open(".apikey")
openai.api_key = file.readline()
file.close()

# list models
models = openai.Model.list()

# print the first model's id
print(models.data[0].id)
