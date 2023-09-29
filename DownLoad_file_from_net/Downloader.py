import os
import shutil
import requests
import zipfile
import pandas as pd

"""
pip install requests
pip install pandas
pip install xlrd
"""

Path = r"total address"#总路径
xls_file_path = os.path.join(Path, "下载目录.xls")
record_zip_Path = os.path.join(Path, "ZIP文件夹")
result_Path = os.path.join(Path, "处理结果")


# 下载网址
url_basic = "your url"

# 使用pandas的read_excel函数读取XLS文件
df = pd.read_excel(xls_file_path)

record_toatal = df.shape[0]  # 获取记录总数目
factory_id = df.loc[:,"factory_id"]

# 打印数据框的内容
# print(df)
# print(record_toatal)

# df_cut = df[1:record_toatal]#去标题



def DownLoad_fun(imgpack_id,filename,download_file_path,output_folder_name):
    # 定义要下载的文件的URL
    url_target = url_basic + imgpack_id
    print("正在从",url_target,"下载...")

    # 发起GET请求下载文件
    response = requests.get(url_target)

    # 检查是否成功下载文件
    if response.status_code == 200:

        print("已成功下载文件")
        # 拼接目标文件夹的完整路径
        output_folder_path = os.path.join(os.getcwd(), output_folder_name)

        # 写入下载的文件
        with open(download_file_path, 'wb') as output_file:
            output_file.write(response.content)


        unzip = os.path.join(Path, "Cache")
        # 解压缩文件到新文件夹
        with zipfile.ZipFile(download_file_path, 'r') as zip_ref:
            zip_ref.extractall(unzip)
        
        print("文件解压完成，已放入文件夹:", unzip)

        # 遍历解压后的文件夹
        for root, dirs, files in os.walk(unzip):
            for file in files:
                file_path = os.path.join(root, file)
                if not file.lower().endswith(".jpg"):
                    # 文件不是 JPG 类型
                    # 删除其他文件
                    os.remove(file_path)
                else:
                    # 文件是 JPG 类型
                    # 获取文件的新名称，例如使用 factory_id + 文件名
                    new_file_name = filename.replace(".zip", "") + "_" + str(int(file.replace("0", "").replace(".jpg", ""))-1)+".jpg"
                    new_file_path = os.path.join(root, new_file_name)
                    # 重命名文件
                    os.rename(file_path, new_file_path)
                    print(f"重命名文件: {file} 为 {new_file_name}")
                    # 复制文件到目标文件夹
                    shutil.copy(new_file_path, output_folder_name)
                    print(f"复制文件：{new_file_name} 到文件夹：{output_folder_name}")
                    # 删除原始文件
                    os.remove(new_file_path)
                    print(f"删除文件：{new_file_name} 成功")


        print("保留.jpg文件操作成功")
    else:
        print("下载文件失败")




os.makedirs(record_zip_Path, exist_ok=True)
factory_list = []#存放工厂id
# 使用iterrows()方法遍历DataFrame的每行数据
for index, row in df.iterrows():

    print("#########################处理开始#########################")
    print("*开始处理第", index, "行数据")
    print("*工厂名称为:", row[3], "*filename为:",row[2])
    # print(row[0])#factory_id
    # print(row[1])#imgpack_id
    # print(row[2])#filename
    # print(row[3])#factory_name


    #处理工厂重复与否，否则要新建文件夹
    if row[0] not in factory_list:
        print("数据解析完成，需要新建文件夹")
        factory_list.append(row[0])
        # 定义目标文件夹的名称
        output_folder_name = os.path.join(result_Path, row[3])
        # 创建目标文件夹
        os.makedirs(output_folder_name, exist_ok=True)
        print("文件夹已新建，准备下载!")
    else:
        print("数据解析完成，已检测到目标文件夹，准备下载!")



    download_file_path = os.path.join(record_zip_Path, row[2])
    output_folder_name = os.path.join(result_Path, row[3])
    DownLoad_fun(str(row[1]),row[2],download_file_path,output_folder_name)


    print("#########################处理结束#########################")

