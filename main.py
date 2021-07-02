import pathlib
import xml.etree.ElementTree
import zipfile


def get_task(tasks):
    alphabet = 'Й Ц У К Е Н Г Ш Щ З Х Ъ Ф Ы В А П Р О Л Д Ж Э Ё Я Ч С М И Т Ь Б Ю' \
               'й ц у к е н г ш щ з х ъ ф ы в а п р о л д ж э ё я ч с м и т ь б ю'
    list_tasks = []
    for task in tasks:
        i = 0
        while i < task.__len__():
            symbol = task[i:i + 1]
            if symbol in alphabet:
                list_tasks.append(task[i:])
                break
            else:
                i += 1
    task_return = []
    for reverse in list_tasks:
        task_return.append(reverse[::-1])

    return task_return


def get_text_in_task(tasks, texts):
    for task in tasks:
        for text in texts:
            if task in text:
                i = texts.index(text)
                task_2 = tasks[(tasks.index(task) + 1)]
                for text_2 in texts:
                    if task_2 in text_2:
                        j = texts.index(text_2)
                text1 = texts[i + 1:j - 1]
                TEXT = ''
                for str_text in text1:
                    TEXT = TEXT + '\n\t' + str_text
                # todo: Вставить проверку задачи в системе и добавить сохранение задачи и текста


def main():
    currentDirectory = pathlib.Path('./files/')

    for currentFile in currentDirectory.iterdir():
        # todo: Добавить создание проекта

        WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        PARA = WORD_NAMESPACE + 'p'
        TEXT = WORD_NAMESPACE + 't'
        list_task = []
        with zipfile.ZipFile(currentFile) as docx:
            tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))

        """Текст лист"""
        list_text = []
        for text in tree.iter(PARA):
            item = ''.join(node.text for node in text.iter(TEXT))
            if item == '' or item == ' ' or item.__len__() < 2:
                continue
            else:
                list_text.append(item)

        for task in list_text:
            if task == 'Конец оглавления':
                break
            else:
                list_task.append(task)

        tasks = get_task(list_task)
        """Конечный список задач"""
        tasks = get_task(tasks)

        texts = list_text[tasks.__len__() + 1:]

        """Получение текста для задачи"""
        get_text_in_task(tasks, texts)


if __name__ == '__main__':
    main()
