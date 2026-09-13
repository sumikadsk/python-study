from openpyxl import load_workbook, Workbook
import os
import pymysql
import redis

# ================= 配置区 =================
MYSQL_CONFIG = {
    'host': '127.0.0.1',
    'user': 'root',          # MySQL账号
    'password': '123456',      # MySQL密码
    'database': 'office_db', # 事先建好的数据库
    'charset': 'utf8mb4'
}
# 连接Redis
r = redis.Redis(host='127.0.0.1', port=6379, db=0, decode_responses=True)

class ExcelAutoHandler:
    """Excel自动化处理工具：数据清洗、去重、统计、自动生成报表 + 数据库持久化 + Redis幂等防重"""

    def __init__(self):
        # 新建最终输出工作簿
        self.wb_out = Workbook()
        self.ws_out = self.wb_out.active
        self.ws_out.title = "数据汇总结果"
        # 标志位：确保表头只写一次
        self.is_header_written = False
        self.is_db_init = False

    def load_excel_data(self, file_path):
        """读取单个Excel文件数据"""
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            data_list = []
            header = None

            # 遍历每行数据
            for idx, row in enumerate(ws.iter_rows(values_only=True)):
                # 过滤空行
                if not any(cell is not None for cell in row):
                    continue
                
                # 提取第一行作为表头
                if idx == 0:
                    header = row
                    continue
                
                data_list.append(row)
            return header, data_list
        except Exception as e:
            print(f"读取文件失败：{file_path}，错误：{e}")
            return None, []

    def clean_data(self, data):
        """数据清洗：去重（保留原有顺序）"""
        # set会打乱顺序，用dict.fromkeys保持顺序去重
        clean_data = list(dict.fromkeys(data))
        return clean_data

    def init_db_table(self, header):
        """【新增】动态初始化MySQL表结构，防止表不存在报错"""
        if self.is_db_init:
            return
        try:
            conn = pymysql.connect(**MYSQL_CONFIG)
            cursor = conn.cursor()
            # 这里为了演示，写死你截图里的表结构；实际可以动态根据header生成
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

    def save_result(self, save_path="汇总报表.xlsx"):
        """自动保存处理后的Excel报表"""
        self.wb_out.save(save_path)
        print(f"自动化处理完成，报表已导出：{save_path}")

    def run_batch(self, folder_path):
        """批量处理文件夹内所有xlsx文件"""
        all_clean_data = []
        
        # 遍历文件夹
        for file in os.listdir(folder_path):
            if file.endswith((".xlsx", ".xls")):
                file_full = os.path.join(folder_path, file)
                file_name = os.path.basename(file_full)

                # ========== 新增：Redis 幂等防重 ==========
                # 处理前，检查Redis里有没有这个文件的处理记录
                if r.get(f"excel_done:{file_name}"):
                    print(f"[跳过] {file_name} 已经处理过了 (Redis防重)")
                    continue
                # =======================================

                # 读取数据
                header, raw_data = self.load_excel_data(file_full)
                if not raw_data:
                    continue
                
                # 初始化数据库
                self.init_db_table(header)

                # 清洗数据
                clean_data = self.clean_data(raw_data)
                
                # 写入Excel（第一行写表头）
                if not self.is_header_written and header:
                    self.ws_out.append(header)
                    self.is_header_written = True
                for row in clean_data:
                    self.ws_out.append(row)

                # ========== 新增：MySQL 数据持久化 ==========
                try:
                    conn = pymysql.connect(**MYSQL_CONFIG)
                    cursor = conn.cursor()
                    # 使用 executemany 批量插入，大幅提高效率
                    sql = "INSERT INTO employee_data (name, department, salary, hire_date) VALUES (%s, %s, %s, %s)"
                    
                    # 确保元组长度匹配（4列）
                    values = [row[:4] for row in clean_data if len(row) >= 4]
                    
                    if values:
                        cursor.executemany(sql, values)
                        conn.commit()
                        print(f"[成功] {file_name} 数据已入库，并在Redis标记完成")
                    
                    # 处理成功后，在 Redis 记录状态，设置7天过期
                    r.setex(f"excel_done:{file_name}", 604800, "done")
                    
                except Exception as e:
                    print(f"[失败] {file_name} 数据库写入出错: {e}")
                finally:
                    if conn:
                        conn.close()
                # ===========================================

        # 保存最终报表
        self.save_result()


if __name__ == "__main__":
    # 改成你存放Excel文件的文件夹路径
    target_folder = r"C:\Users\13770\Desktop\excel_data"

    # 实例化并执行自动化流程
    tool = ExcelAutoHandler()
    tool.run_batch(target_folder)
