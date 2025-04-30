import os
import pandas as pd
from pymongo import MongoClient
from pymongo.errors import ConnectionFailure
from typing import List
import zipfile
import tempfile
import shutil

# 定义数据格式异常类
class UnrecognizedDataFormatError(Exception):
    pass

class DataLoader:
    def __init__(self, config: dict):
        # 新增线索数据暂存列表
        self.clue_dfs = []  # 用于暂存需要延迟处理的线索数据
        # 增强路径校验逻辑
        self.file_path = config.get('data_source', {}).get('file_dir')
        if not self.file_path:
            raise ValueError("配置文件中缺失数据源目录路径(data_source.file_dir)")
            
        # 将相对路径转换为绝对路径
        if not os.path.isabs(self.file_path):
            self.file_path = os.path.join(os.getcwd(), self.file_path)
            
        if not os.path.exists(self.file_path):
            raise ValueError(f"数据源目录不存在，请检查配置: {self.file_path}\n建议：1.创建目录 2.在config.yaml中配置正确路径")
        
        self.mongo_uri = config.get('database', {}).get('mongo_uri')
        self.db_name = config.get('database', {}).get('db_name')
        
        if not self.mongo_uri or not self.db_name:
            raise ValueError("MongoDB配置不完整")
        
        self.client = MongoClient(self.mongo_uri)
        self.db = self.client[self.db_name]

        # 新增集合存在性检查逻辑
        self._ensure_collections_exist()
        
        # 修正字段模板为实际存在的中文字段
        self.clue_fields = {'本端身份证号', '本端预留手机号', '交易金额', '交易时间', '本端账户名'}
        self.bank_account_fields = {'账户', '账户名', '余额', '开户行', '开户日期'}  # 原字段含'本端账户'等前缀
        self.transaction_fields = {'交易金额', '交易时间', '交易类型', '对端账户名', '摘要', '交易网点'}

        # 调整匹配规则顺序和最小匹配数（修复线索误识别问题）
        self.matching_rules = [
            (self.clue_fields, 'sales_clue', 3),  # 优先处理线索类型
            (self.bank_account_fields, 'bank_account', 3),
            (self.transaction_fields, 'bank_transaction', 4)  # 提升交易匹配阈值
        ]

    def check_mongo_connection(self):
        """检测MongoDB连接是否正常"""
        try:
            # ping命令用于检测连接状态
            self.client.admin.command('ping')
            print("MongoDB连接正常")
        except ConnectionFailure as e:
            raise ConnectionFailure(f"MongoDB连接失败: {str(e)}")
        except Exception as e:
            raise Exception(f"数据库连接异常: {str(e)}")

    def load_files(self):
        """加载目录下的所有数据文件"""
        print(f"正在加载目录: {self.file_path}")
        files_loaded = False

        # 新增ZIP文件处理逻辑
        zip_files = [f for f in os.listdir(self.file_path) if f.lower().endswith('.zip')]
        for zip_file in zip_files:
            zip_path = os.path.join(self.file_path, zip_file)
            print(f"\n发现ZIP压缩文件: {zip_file}，开始解压...")
            temp_dir = tempfile.mkdtemp()
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                    print(f"解压完成，临时目录: {temp_dir}")

                # 处理解压后的文件
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        if file.lower().endswith(('.xlsx', '.xls', '.csv')):
                            print(f"处理解压文件: {file}")
                            self.process_file(file_path)
                            files_loaded = True
            except zipfile.BadZipFile as e:
                print(f"ZIP文件损坏，跳过处理: {zip_file}，错误: {str(e)}")
            finally:
                shutil.rmtree(temp_dir)

        # 第一阶段：优先处理非线索文件
        for file in os.listdir(self.file_path):
            if file.endswith(('.xlsx', '.xls', '.csv')):
                file_path = os.path.join(self.file_path, file)
                print(f"处理非线索文件: {file}")
                self.process_file(file_path)
                files_loaded = True
        
        # 第二阶段：最后处理线索文件
        if self.clue_dfs:
            print("\n=== 开始处理暂存的线索数据 ===")
            total_clues = sum(len(df) for df in self.clue_dfs)
            print(f"▷ 共发现 {total_clues} 条待处理线索")
            for idx, df in enumerate(self.clue_dfs, 1):
                self._annotate_transaction(df)
                print(f"▏ 进度: [{idx}/{len(self.clue_dfs)}] 个线索文件处理完成")
        
        if not files_loaded:
            print("未找到任何数据文件")

    def process_file(self, file_path: str):
        """处理单个数据文件"""
        try:
            # 新增文件名关键词检测（优先处理线索文件）
            file_name = os.path.basename(file_path)
            if "线索" in file_name.lower():
                df = pd.read_excel(file_path, engine='openpyxl') if file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(file_path)
                self.clue_dfs.append(df)
                print(f"▶ 发现线索文件（文件名识别）: {file_name}，已暂存等待后续处理")
                return  # 直接返回，跳过后续字段匹配流程

            if file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path, engine='openpyxl')  # 明确指定引擎
            else:
                df = pd.read_csv(file_path)
            
            # 增强列名检测逻辑
            actual_columns = set(df.columns.str.strip())
            
            # 添加列名统一处理逻辑
            normalized_columns = {col.replace('(组)', '').replace('（组）', '').strip() for col in actual_columns}
            
            # 初始化匹配追踪变量（修复变量未初始化问题）
            max_matches = 0
            best_match = None
            
            for fields, collection, min_match in self.matching_rules:
                matched_fields = normalized_columns & fields
                # 同时满足最小匹配数和最大匹配数要求
                if len(matched_fields) >= min_match and len(matched_fields) > max_matches:
                    max_matches = len(matched_fields)
                    best_match = collection
            
            if best_match:
                if best_match == 'sales_clue':
                    # 线索数据暂存延迟处理
                    self.clue_dfs.append(df)
                    print(f"▶ 发现线索文件（字段匹配）: {file_path}，已暂存等待后续处理")
                else:
                    # 新增账户数据识别日志
                    if best_match == 'bank_account':
                        matched_fields = normalized_columns & self.bank_account_fields
                        print(f"✔ 检测到账户数据文件: {file_name}")
                        print(f"   匹配字段: {', '.join(matched_fields)}")
                    self._save_to_mongo(df, best_match)
            else:
                # 生成更友好的错误提示，列出所有模板要求
                expected_fields = "\n".join(
                    [f"- {collection}需要字段: {', '.join(fields)}" 
                     for fields, collection, _ in self.matching_rules]
                )
                raise UnrecognizedDataFormatError(
                    f"无法识别文件格式: {file_path}\n"
                    f"实际列名: {list(df.columns)}\n"
                    f"支持的字段组合:\n{expected_fields}"
                )

        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {str(e)}")

    def _annotate_transaction(self, clue_df: pd.DataFrame):
        """标注异常交易数据"""
        transaction_col = self.db['bank_transaction']
        annotated_count = 0
        supplemented_count = 0
        
        # 新增：批量收集需要补充的数据
        supplement_data = []
        
        # 查找匹配的交易记录进行标注
        for _, row in clue_df.iterrows():
            # 根据关键字段匹配交易数据
            match_filter = {
                '交易时间': row.get('交易时间'),
                '交易金额': row.get('交易金额'),
                '本端账户名': row.get('本端账户名')
            }
            
            # 更新交易记录添加异常标记（增加未标记条件）
            update_result = transaction_col.update_many(
                {'$and': [match_filter, {'异常标记': {'$ne': True}}]},  # 新增条件：仅更新未标记异常的记录
                {'$set': {'异常标记': True, '线索详情': row.to_dict()}}
            )
            
            # 记录标注数量
            if update_result.matched_count > 0:
                annotated_count += update_result.matched_count
                print(f"✓ 标注异常交易 [{update_result.matched_count}条] 时间:{row.get('交易时间')} 金额:{row.get('交易金额')}")
            elif transaction_col.count_documents(match_filter) > 0:  # 新增已标记检查
                print(f"ⓘ 交易已标记异常，跳过处理 时间:{row.get('交易时间')} 金额:{row.get('交易金额')}")

        # 新增：批量存储补充数据
        if supplement_data:
            self._save_to_mongo(pd.DataFrame(supplement_data), 'bank_transaction')
            
        # 添加统计日志
        print(f"▌ 处理完成 - 异常标注: {annotated_count}条 | 补充数据: {supplemented_count}条\n")

    def _save_to_mongo(self, df: pd.DataFrame, collection_name: str):
        """将数据存入MongoDB"""
        try:
            collection = self.db[collection_name]
            records = df.to_dict('records')
            inserted_count = 0
            duplicate_count = 0
            
            for record in records:
                if collection_name == 'bank_transaction':
                    # 新增：设置异常标记默认值为False
                    record.setdefault('异常标记', False)
                    # 修复：为交易记录添加查询条件
                    query = {
                        '交易时间': record.get('交易时间'),
                        '交易金额': record.get('交易金额'),
                        '本端账户名': record.get('本端账户名'),
                        '对端账户名': record.get('对端账户名')
                    }
                else:
                    # 修复查询条件缺失问题
                    if collection_name == 'bank_account':
                        query = {'account': record.get('account')}  # 使用账户号作为唯一标识
                    elif collection_name == 'sales_clue':
                        query = {'线索ID': record.get('线索ID')}  # 假设存在唯一标识字段
                    else:
                        query = record.copy()  # 兜底方案

                # 检查是否存在重复记录
                if collection.find_one(query):
                    print(f"ⓘ 数据已存在，跳过: {collection_name} {query}")
                    duplicate_count += 1
                else:
                    collection.insert_one(record)
                    inserted_count += 1

            print(f"▣ 存储完成 - 集合: {collection_name} | 新增: {inserted_count}条 | 跳过: {duplicate_count}条")
            # 增强账户数据存储反馈
            if collection_name == 'bank_account' and inserted_count > 0:
                print(f"█ 成功存入{inserted_count}条账户数据到{collection_name}集合")
        except Exception as e:
            print(f"插入数据到 {collection_name} 时出错: {str(e)}")

    def _upsert_account(self, account_data: dict):
        """更新或插入账户数据（统一处理本端/对端）"""
        if not account_data.get('account'):
            return
            
        collection = self.db['bank_account']
        query = {'account': account_data['account']}
        
        # 设置默认值（新增余额字段默认值）
        account_data.setdefault('balance', 0.0)
        account_data.setdefault('is_abnormal', False)  # 确保默认值为False
        
        # 统一字段命名规范（删除本端/对端前缀）
        standardized_data = {
            'account': account_data['account'],
            'account_name': account_data.get('account_name', ''),
            'bank': account_data.get('bank', '未知支行'),
            'open_date': account_data.get('open_date', '1970-01-01'),
            'balance': account_data['balance']
        }
        
        # 检查关联交易的异常状态（增强关联检测逻辑）
        abnormal_count = self.db['bank_transaction'].count_documents({
            '$or': [
                {'本端账户名': standardized_data['account_name']},
                {'对端账户名': standardized_data['account_name']}
            ],
            '异常标记': True
        })
        standardized_data['is_abnormal'] = abnormal_count > 0  # 自动更新异常状态

        # 仅当发现新异常时才更新
        if not standardized_data['is_abnormal'] and abnormal_count > 0:
            standardized_data['is_abnormal'] = True

        # 修复更新冲突问题
        update_operation = {
            '$setOnInsert': {k:v for k,v in standardized_data.items() if k != 'is_abnormal'},
            '$set': {'is_abnormal': standardized_data['is_abnormal']}
        }
        
        result = collection.update_one(
            query,
            update_operation,
            upsert=True
        )
        
        if result.upserted_id:
            print(f"  新增账户数据 - 账号: {standardized_data['account']} 户名: {standardized_data['account_name']}")
        elif result.modified_count > 0:
            print(f"  更新异常状态 - 账号: {standardized_data['account']} 异常: {standardized_data['is_abnormal']}")

    def _ensure_collections_exist(self):
        """确保所需的数据库集合存在"""
        required_collections = ['sales_clue', 'bank_account', 'bank_transaction']
        existing_collections = self.db.list_collection_names()
        
        for coll in required_collections:
            if coll not in existing_collections:
                self.db.create_collection(coll)
                print(f"▧ 已创建集合: {coll}")