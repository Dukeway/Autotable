import requests
import json
import logging
from ollama import Client

logger = logging.getLogger(__name__)

class BaseLLMClient:
    """大语言模型客户端基类"""
    def chat_completion(self, messages, temperature=0.7):
        raise NotImplementedError("子类必须实现chat_completion方法")

class APIClient(BaseLLMClient):
    """OpenAI兼容API客户端"""
    def __init__(self, api_base_url, api_key, model_name):
        self.api_base_url = api_base_url.rstrip("/")
        self.api_key = api_key
        self.model_name = model_name
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

    def chat_completion(self, messages, temperature=0.7):
        url = f"{self.api_base_url}/chat/completions"
        payload = {
            "model": self.model_name,
            "messages": messages,
            "temperature": temperature
        }
        try:
            response = requests.post(url, headers=self.headers, json=payload)
            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]
        except requests.exceptions.RequestException as e:
            logger.error(f"API请求失败: {str(e)}")
            raise
        except KeyError as e:
            logger.error(f"API响应格式异常: {str(e)}")
            raise

class OllamaClient(BaseLLMClient):
    """本地Ollama客户端"""
    def __init__(self, host, model_name):
        self.client = Client(host=host)
        self.model_name = model_name

    def chat_completion(self, messages, temperature=0.3):
        try:
            response = self.client.chat(
                model=self.model_name,
                messages=messages,
                options={"temperature": temperature}
            )
            return response['message']['content']
        except Exception as e:
            logger.error(f"Ollama调用失败: {str(e)}")
            raise