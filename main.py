from core.data_loader import DataLoader
import yaml
import os

# 定义自定义异常类
class UnrecognizedDataFormatError(Exception):
    """自定义异常类，用于处理数据格式错误"""
    pass

def main(config_path):
    # 读取配置文件
    if not os.path.exists(config_path):
        print(f"配置文件不存在: {config_path}")
        return
    
    with open(config_path, 'r') as file:
        config = yaml.safe_load(file)
    
    # 初始化数据加载器
    try:
        loader = DataLoader(config)
    except Exception as e:
        print(f"初始化数据加载器失败: {str(e)}")
        return
    
    # 新增MongoDB连接检测
    try:
        loader.check_mongo_connection()
    except Exception as e:
        print(f"\n错误: {str(e)}\n请检查MongoDB服务状态和配置")
        return
    
    # 加载数据文件
    try:
        loader.load_files()
        print("数据加载完成")
    except UnrecognizedDataFormatError as e:
        print(f"\n数据格式错误: {str(e)}\n请检查文件字段是否符合模板要求")
    except Exception as e:
        print(f"数据加载失败: {str(e)}")

if __name__ == "__main__":
    # 使用相对路径加载配置文件
    config_path = os.path.join(os.path.dirname(__file__), 'config', 'config.yaml')
    main(config_path)