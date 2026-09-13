from openpyxl import load_workbook, Workbook
import os
import pymysql
import redis

# ================= 配置区 =================
MYSQL_CONFIG = {
    'host': '127.0.0.1',
    'user': 'root',
    'password': '123456',      # 确认这是你本地的MySQL密码
    'database': 'office_db',
    'charset': 'utf8mb4'
}
# 连接Redis
r = redis.Redis(host='127.0.0.1', port=6379, db=0, decode_responses=True)

class ExcelAutoHandler:
    """Excel自动化处理工具：数据清洗、去重、统计、自动生成报表 + 数据库持久化 + Redis幂等防重"""

    def __init__(self):
        self.wb_out = Workbook()
        self.ws_out = self.wb_out.active
        self.ws_out.title = "数据汇总结果"
        self.is_header_written = False
        self.is_db_init = False

    def load_excel_data(self, file_path):
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            data_list = []
            header = None

            for idx, row in enumerate(ws.iter_rows(values_only=True)):
                if not any(cell is not None for cell in row):
                    continue
                if idx == 0:
                    header = row
                    continue
                data_list.append(row)
            return header, data_list
        except Exception as e:
            print(f"读取文件失败：{file_path}，错误：{e}")
            return None, []

    def clean_data(self, data):
        return list(dict.fromkeys(data))

    def init_db_table(self, header):
        if self.is_db_init:
            return
        try:
            conn = pymysql.connect(**MYSQL_CONFIG)
            cursor = conn.cursor()
            create_table_sql = """
            CREATE TABLE IF NOT EXISTS employee_data (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(50),
                department VARCHAR(50),
                salary DECIMAL(10, 2),
                hire_date VARCHAR(20)
            );
            """
            cursor.execute(create_table_sql)
            conn.commit()
            conn.close()
            self.is_db_init = True
        except Exception as e:
            print(f"数据库初始化失败，请检查MySQL是否启动: {e}")

    def save_result(self, save_path):
        try:
            self.wb_out.save(save_path)
            print(f"自动化处理完成，报表已导出到：{save_path}")
        except PermissionError:
            print(f"导出失败！请先关闭正在打开的 {save_path}，然后再运行一次！")
        except Exception as e:
            print(f"导出失败，错误：{e}")

    def run_batch(self, folder_path):
        for file in os.listdir(folder_path):
            if file.startswith('~$'):
                continue
                
            if file.endswith((".xlsx", ".xls")):
                file_full = os.path.join(folder_path, file)
                file_name = os.path.basename(file_full)

                # Redis 幂等防重
                if r.get(f"excel_done:{file_name}"):
                    print(f"[跳过] {file_name} 已经处理过了 (Redis防重)")
                    continue

                header, raw_data = self.load_excel_data(file_full)
                if not raw_data:
                    continue
                
                self.init_db_table(header)
                clean_data = self.clean_data(raw_data)
                
                # 写入Excel
                if not self.is_header_written and header:
                    self.ws_out.append(header)
                    self.is_header_written = True
                for row in clean_data:
                    self.ws_out.append(row)

                #  MySQL 数据持久化 
                conn = None
                try:
                    conn = pymysql.connect(**MYSQL_CONFIG)
                    cursor = conn.cursor()
                    sql = "INSERT INTO employee_data (name, department, salary, hire_date) VALUES (%s, %s, %s, %s)"
                    
                    values = []
                    for row in clean_data:
                        if len(row) >= 4:
                            try:
                                float(row[2]) 
                                values.append(row[:4])
                            except (ValueError, TypeError):
                                print(f"  -> [警告] 发现脏数据（工资非数字），已跳过: {row}")
                    
                    if values:
                        cursor.executemany(sql, values)
                        conn.commit()
                        print(f"[成功] {file_name} 数据已入库，并在Redis标记完成")
                    
                    r.setex(f"excel_done:{file_name}", 604800, "done")
                    
                except Exception as e:
                    print(f"[失败] {file_name} 数据库写入出错: {e}")
                    if conn:
                        conn.rollback()
                finally:
                    if conn:
                        conn.close()

        # 保存到你的电脑桌面
        desktop_path = r"C:\Users\13770\Desktop\汇总报表.xlsx"
        self.save_result(desktop_path)


#  执行入口 
if __name__ == "__main__":
    # 清空Redis缓存，防止之前测试过的文件被永久跳过
    r.flushdb()
    print("已清空Redis缓存，准备重新处理所有文件...")

    target_folder = r"C:\Users\13770\Desktop\excel_data"
    
    tool = ExcelAutoHandler()
    tool.run_batch(target_folder)
