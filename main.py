import openpyxl
from translate import Translator
from genanki import Model, Deck, Package, Note
from datetime import datetime
from eng_to_ipa import convert

# 定义 Anki 卡片模型
my_model = Model(
    1607392319,
    'Simple Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Phonetic'},
        {'name': 'Translation'},
        {'name': 'Answer'},
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '<div style="font-size: 30px; font-family: Arial; text-align: center; color: black;">{{Question}}</div>',
            'afmt': '{{FrontSide}}<hr id="answer"><div style="font-size: 20px; font-family: Arial; text-align: center; color: black;"><br><span style="color: blue;">/{{Phonetic}}/</span><br><br>Google 翻译: <span style="color: green;">{{Translation}}</span><br><br>DeepL 翻译: <span style="color: purple;">{{Answer}}</span></div>',
        },
    ],
    css='.card { background-color: white; } hr { border-top: 1px solid black; }')

# 创建一个新的 Deck
my_deck = Deck(2059400110, 'Bob Translated Words')

# 读取 Excel 文件
workbook = openpyxl.load_workbook('bob.xlsx', read_only=True)
worksheet = workbook.active

# 创建翻译器对象
translator = Translator(to_lang="zh", provider="mymemory")

# 遍历 Excel 行并创建 Note 对象
for row in worksheet.iter_rows(min_row=2, values_only=True):
    original_text, translated_text = row[4], row[8]

    if original_text and translated_text:
        # 获取单词的 IPA 音标
        phonetic = convert(original_text)

        # 获取单词的翻译
        translation = translator.translate(original_text)

        note = Note(model=my_model, fields=[original_text, phonetic, translation, translated_text])
        my_deck.add_note(note)

# 获取当前时间并格式化为字符串
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

# 生成文件名
output_file = f"bob_{current_time}.apkg"

# 将 Deck 打包为 Anki 包文件
Package(my_deck).write_to_file(output_file)

print(f"Anki 卡片包已生成: {output_file}")