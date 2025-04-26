from watcher.excel_file_watcher import ExcelFileWatcher
from excel_handler.processor import ExcelProcessor
from web_automation.automator import AeonUploader
from excel_handler.workflow import validate_excel_data, create_upload_data_from_processor, match_and_fill_from_csv,move_csv_to_folder,tantou_name,save_excel_object
from ledger.log import log_process_result

# 第一步：监视文件夹创建文件夹
watcher = ExcelFileWatcher()
print("📂 正在持续监听文件夹...")

while True:
    new_file_path, new_folder_path = watcher.wait_for_new_file()
    print("✅ 检测到并移动了文件:", new_file_path)

    a = ExcelProcessor(new_file_path)
    print("📄 正在处理文件:", new_file_path)
    try:
        # 第二步：校验excel数据，返回错误信息
        errors = validate_excel_data(a)
        name = tantou_name(a)
        print("📄 名称:", name)
        if errors:
            print("❌ 校验失败，原因：", errors)
        else:
            # 第三步：生成nagashikomi数据
            save_path = create_upload_data_from_processor(a, new_folder_path)
            print("✅ 上传数据生成完毕", save_path)

            # 第四步：上传数据到 Web，下载CSV
            uploader = AeonUploader()
            result = uploader.run(save_path)
            if result["success"]:
                print("✅ 上传成功")
                # 第五步：从 CSV 匹配并填充，将CSV移动到指定文件夹
                csv_path = match_and_fill_from_csv(processor=a)
                save_path = save_excel_object(processor=a)# 发注书另存为 new 文件
                print(f"💾 文件已保存: {save_path}")
                new_csv_path = move_csv_to_folder(csv_path, new_folder_path)#list移动到new_folder_path之下
                print(f"{csv_path} 已成功移动到 {new_folder_path}")
            else:
                print("❌ 上传失败，原因：", result["error"])

    except Exception as e:
        print("⚠️ 处理流程出错:", str(e))

    finally:
        a.close()
        # 第六步：记录处理结果
        log_process_result(
            log_path="log.csv",
            new_file_path=new_file_path,
            new_folder_path=new_folder_path,
            save_path=locals().get("save_path"),
            name=locals().get("name"),
            errors=locals().get("errors"),
            result=locals().get("result"),
            new_csv_path = locals().get("new_csv_path"),
        )
        print("📄 文件处理完毕，继续监听中...\n")
