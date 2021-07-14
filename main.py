import excel_worker as xlwork
import json


def main():
    path = input('Введите путь к файлу: ')
    json_out = xlwork.get_json(path)

    with open('result.json', 'w', encoding='utf-8') as fp:
        json.dump(json_out, fp, indent=4, ensure_ascii=False)
        print('JSON-файл готов')


if __name__ == '__main__':
    main()