import logging
from config import RUN_MODE, API_BASE_URL, API_KEY, API_MODEL_NAME, OLLAMA_HOST, OLLAMA_MODEL_NAME
from llm_clients import APIClient, OllamaClient
from autotable import AutoTable

# 配置日志系统
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("autotable.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

if __name__ == "__main__":
    if RUN_MODE == "api":
        logger.info("使用 API 模式")
        llm_client = APIClient(
            api_base_url=API_BASE_URL,
            api_key=API_KEY,
            model_name=API_MODEL_NAME
        )
    elif RUN_MODE == "ollama":
        logger.info("使用 Ollama 模式")
        llm_client = OllamaClient(
            host=OLLAMA_HOST,
            model_name=OLLAMA_MODEL_NAME
        )
    else:
        raise ValueError(f"无效的运行模式: {RUN_MODE}")

    auto_table = AutoTable(
        knowledge_base_path="知识库.xlsx",
        word_template_path="表格模板.docx",
        llm_client=llm_client
    )

    auto_table.run()