# _*_ coding:utf-8 _*_
import os

import xlrd


def convert(path):
    """
    将给定的excel文件转换为android 资源文件xml
    :param path: 文件路径，如果与脚步在同一目录下则直接为文件名，如果与脚步不在同一目录下则需要绝对路径
    :return: 成功为True,失败False，以及失败信息；转换为对应的xml输出到当前脚步目录下。
    """

    try:
        data = xlrd.open_workbook(path)
    except FileNotFoundError:
        return False, '文件不存在，请确实后重新输入！'
    table = data.sheets()[0]  # 打开第一张表
    nrows = table.nrows
    if nrows < 2:
        return False, '行数少于2行，不符合转换条件！'
    # 表的有多少列
    ncols = table.ncols
    if ncols < 2:
        return False, '少于两列，不符合转换条件！'
    # 在内存中创建ncols - 1 个空的文档
    docs = []  # [xml.dom.minidom.Document()] * (ncols - 1)
    import xml.dom.minidom
    for i in range(ncols - 1):
        docs.append(xml.dom.minidom.Document())
    # 为每个文档创建resources根节点
    for doc in docs:
        root = doc.createElement('resources')
        doc.appendChild(root)
    names = []
    for i in range(nrows):
        if i == 0:
            names = table.row_values(i)[1:]
            continue
        items = table.row_values(i)[:]
        for n in range(1, len(items)):
            doc = docs[n - 1]
            node = doc.createElement('string')
            node.setAttribute('name', items[0])
            node.appendChild(doc.createTextNode(items[n]))
            root = doc.childNodes[0]
            root.appendChild(node)

    # 写xml文档
    for i in range(len(names)):
        xml = open('strings_%s.xml' % names[i], 'w')
        docs[i].writexml(xml, indent='\n', addindent='\t', newl='\n', encoding='utf-8')
        xml.close()
    return True, 'success'


if __name__ == '__main__':
    print('=' * 8, "使用说明", '=' * 8)
    print("""
    *  功能说明：
    *      该脚步为将如下格式的excel文件转换为android多语言xml文件
    *  | name | language_1 | language_2 |
    *  其中：
    *      name 对应<string name='xxx'> 中的xxx，
    *      language_x: 对应不同语言下xxx的值，如下列子：
    *      |name    |zh  |en   |
    *      |app_name|微信|WeChat|
    *  输出：strings_language_x.xml
    """)
    print('=' * 8, "  end  ", '=' * 8)
    print('请输入要转换文件的文件名，注意如果不是与该脚本在同一目录下则输入绝对路径：')
    while True:
        path = input('--->')
        if not path:
            print('输入路径不能为空，请重新输入！如果要退出，直接ctrl+c或者输入exit')
        elif path is 'exit':
            break
        else:
            print("正在转换处理，请稍后。。。")
            success, errormsg = convert(path)
            if success:
                print('恭喜，转换成功！！！')
                break
            else:
                print('转换失败，', errormsg)