#!/usr/bin/python
# -*- coding: UTF-8 -*-
import random
import genanki

# 生成随机Note GUID
class MyNote(genanki.Note):
  @property
  def guid(self):
    return genanki.guid_for(self.fields[0], self.fields[1])

'''
#根据字符串生成随机数
def random_20char(string,length):
    for i in range(length):
        x = random.randint(1,2)
        if x == 1:
            y = str(random.randint(0,9))
        else:
            y = chr(random.randint(97, 122))
        string.append(y)
    string = ''.join(string)
    return string
'''


# anki操作
class anki_generation(object):
    # model id与name, deck id与name
    def __init__(self, model_name=None, deck_name=None):
        #  初始化model
        self.model_id = random.randrange(1 << 30, 1 << 31)  # 随机model_id
        self.deck_id = random.randrange(1 << 29, 1 << 32)  # 随机deck_id

        self.my_model = genanki.Model(
            self.model_id,
            model_name,
            fields=[
                {'name': 'chinese'},
                {'name': 'de_word_gender'},
                {'name': 'de_word'},
                {'name': 'sample'},
                {'name': 'MyMedia'},
            ],
            templates=[
                {
                    'name': 'Deutschland Word',
                    'qfmt': '{{chinese}}',
                    'afmt': '{{FrontSide}}<hr id="answer">{{de_word_gender}}</br>{{de_word}}</br>{{sample}}</br>{{MyMedia}}',
                },
        ])
        # 初始化deck
        self.my_deck = genanki.Deck(
            self.deck_id,
            deck_name)


    # 添加note
    def add_word_into_note(self, gender, word, chinese, example, sound):
        self.my_note = genanki.Note(
            model=my_model,
            fields=[gender, word, chinese, example, sound])
        
        self.my_deck.add_note(self.my_note)


    # 输出apkg
    def output_pk(self):
        outname = deck_name + '.apkg'
        genanki.Package(self.my_deck).write_to_file(outname)


# 主函数
if __name__ == '__main__':
    #  全局id,不同包需要变化
    model_id = 1275934744
    deck_id = 2043828970
    model_name='Deutschland'
    deck_name='Lektion 15_part1'