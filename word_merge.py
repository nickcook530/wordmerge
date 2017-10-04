from mailmerge import MailMerge
#import os

template = input('Please enter your template name:\n')+'.docx'
document = MailMerge(template)

fields = document.get_merge_fields()
MyDict = {}

def menu_loop():
    print('These are you available fields:\n{}'.format(fields))
    loop = True
    while loop:
        #os.system('cls')
        print('Enter z to save and finish')
        answer = input('Enter Placeholder Word = New Word:\n')
        if answer.lower().strip() == 'z':
            print('Are you sure you want to quit? You have incomplete fields!')
            final_answer = input('Enter y or n:\n')
            if final_answer.lower().strip() == 'y':
                loop = False
            else:
                continue
        else:
            fw = first_word(answer)
            if fw in fields:
                sw = second_word(answer)
                MyDict[fw] = sw
                fields.remove(fw)
                if fields:
                    print('Remaining available fields:\n{}'.format(fields))
                else:
                    print('No fields remaining, time to move on')
                    loop = False
            else:
                print('Sorry, that word is not in your merge field!')

def first_word(s):
    holder = "".join(s.split())
    holder = holder.split('=')[0].upper()
    return holder

def second_word(s):
    replacement = s.split('=')[1].strip()
    return replacement

def merge():
    name = input('Enter a name for your new document:\n')
    name = name+'.docx'
    document.merge(**MyDict)
    document.write(name)
    return print('ALL DONE!')

if __name__ == '__main__':
    menu_loop()
    merge()