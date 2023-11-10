import os
import shutil
import re
def extract_number(s):
    pattern = "202\d{9}"
    match = re.search(pattern, s)
    if match:
        return match.group(0)
    else:
        return False
def rename_files(folder_path, output_folder,search_strings,houzhui_string):
    # 遍历文件夹及其子文件夹中的所有文件
    for root, dirs, files in os.walk(folder_path):
        for filename in files:

            if filename.endswith(('.doc', '.docx')):
                # 构建原始文件的完整路径
                file_path = os.path.join(root, filename)

                # 1如果file_path整个的目录地址中出现过“实验4”、“实验四”,然后根据这个来抽出来
                # 1.1 使用 replace() 方法将斜杠字符替换成空格
                modified_file_path_with_space = file_path.replace('\\', ' ')

                # 1.2 当最深层文件夹下的文件为doc或者docx文档格式时，搜索“实验1”、“实验一”而没有出现”操作题“的那些文档
                # 要搜索的子字符串列表
                # 

                # 要排除的子字符串列表
                exclude_strings = ['操作题','操作题1文档.doc','操作题1文档.docx','操作题2文档.doc','操作题2文档.docx','操作题3文档.doc','操作题3文档.docx','操作题4文档.doc','操作题4文档.docx']
                # 初始化标志
                found_once_in_search_strings = False
                # 遍历子字符串列表
                for substring in search_strings:
                    if substring in modified_file_path_with_space:
                        found_once_in_search_strings = True
                        break
                
                found_excluded_strings = any(exclude_string in modified_file_path_with_space for exclude_string in exclude_strings)
                # 检查标志 
                if found_once_in_search_strings and found_excluded_strings==False :
                    # print(modified_string)
                    # print("字符串中包含需要的子字符串，并且不包含任何不需要的子字符串。")

                    # 1.3 将这些文档命名为学号空格实验.doc复制到1word/new文件夹下的班级/实验4/下
                    # 定义要匹配的特定数字
                    # target_pattern = r'\b202\d{9}\b'  # 匹配学号
                    # 特殊情况匹配路径中带中文的情况,不考虑正则匹配中文，考虑切分学号姓名连载一起的
                    # target_pattern = r'\b202\d{9}\b' 
                    # modified_file_path_with_space_all = modified_file_path_with_space.replace(' ','').split(r'\d{12}')
                    # # 使用正则表达式搜索匹配的内容
                    # matches = re.findall(target_pattern, modified_file_path_with_space_all )

                    matchess = extract_number(modified_file_path_with_space)
                    if matchess:
                        # modified_file_path_with_space.split(' ')[-1].split('.')[1]表示文件的后缀名
                        pnew_filename =  os.path.join(matchess+' '+houzhui_string+'.'+modified_file_path_with_space.split(' ')[-1].split('.')[1] )
                        # 构建新文件的完整路径
                        new_file_path = os.path.join(output_folder, pnew_filename)
                        # 将文件复制到指定文件夹并更名
                        shutil.copy(file_path, new_file_path)
                    else:
                        print("未找到匹配到学号 需特殊处理 目录为："+file_path)
                    
                elif found_once_in_search_strings==False and found_excluded_strings==False:
                    # 出现未匹配到实验几也没有匹配到了”操作题“这种情况就要特殊处理，看是不是只有学号命名实验报告的，或者没解压的
                        print("未匹配到实验几也没有匹配到了操作题 需特殊处理 目录为："+file_path)
                else:
                    # # 这种情况就是操作题那种doc不用处理的
                    print("")
                    # print("字符串中未满足搜索条件。")
                

if __name__ == '__main__':
    # 输入的文件夹路径
    folder_path = r'D:\Destop\办公\办公\data\2pdf\机器人21-1'
    # 输出的文件夹路径
    output_folder = r'D:\Destop\办公\办公\data\1word\机器人原21-1\实验4\\'
    # 
    search_strings = ['实验七', '实验7']
    houzhui_string='实验4'

    # 创建输出文件夹（如果不存在）
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 执行文件重命名和复制
    rename_files(folder_path, output_folder,search_strings,houzhui_string)
