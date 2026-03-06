import json
import os
import re
import sys
import shutil


def get_app_dir():
    """
    获取程序实际运行目录：
    - 开发环境：脚本所在目录
    - 打包后：exe 所在目录
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundle_dir():
    """
    获取程序内置资源目录：
    - 开发环境：脚本所在目录
    - PyInstaller 打包后：临时解压目录 _MEIPASS
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


class OfflineTranslator:
    def __init__(self, phrases_file="phrases.json", lexicon_file="lexicon.json"):
        self.app_dir = get_app_dir()
        self.bundle_dir = get_bundle_dir()

        self.phrases_path = self._ensure_external_file(phrases_file)
        self.lexicon_path = self._ensure_external_file(lexicon_file)

        self.phrases = self._load_json(self.phrases_path)
        self.lexicon = self._load_json(self.lexicon_path)

        self.max_phrase_len = self._calc_max_phrase_length()

    def _ensure_external_file(self, filename):
        """
        优先使用程序目录下的外部配置文件。
        如果不存在，则从内置资源复制一份到程序目录。
        """
        external_path = os.path.join(self.app_dir, filename)
        bundled_path = os.path.join(self.bundle_dir, filename)

        if os.path.exists(external_path):
            return external_path

        if os.path.exists(bundled_path):
            shutil.copyfile(bundled_path, external_path)
            return external_path

        raise FileNotFoundError(f"找不到配置文件: {filename}")

    def _load_json(self, file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError(f"词典文件格式错误，必须是 JSON object: {file_path}")

        return data

    def _calc_max_phrase_length(self):
        max_len = 1
        for k in self.phrases.keys():
            word_count = len(k.strip().split())
            if word_count > max_len:
                max_len = word_count
        return max_len

    def translate(self, text):
        if not text or not text.strip():
            return text

        if self._contains_chinese(text):
            return text

        protected_map, protected_text = self._protect_special_tokens(text)
        translated = self._translate_core(protected_text)
        translated = self._restore_special_tokens(translated, protected_map)
        translated = self._cleanup_spaces(translated)
        return translated

    def _translate_core(self, text):
        tokens = self._tokenize(text)
        if not tokens:
            return text

        result = []
        i = 0

        while i < len(tokens):
            token = tokens[i]

            if token["type"] != "word":
                result.append(token["text"])
                i += 1
                continue

            matched = False

            for length in range(min(self.max_phrase_len, len(tokens) - i), 1, -1):
                chunk = tokens[i:i + length]

                if any(t["type"] != "word" for t in chunk):
                    continue

                phrase = " ".join(t["norm"] for t in chunk)
                if phrase in self.phrases:
                    result.append(self.phrases[phrase])
                    i += length
                    matched = True
                    break

            if matched:
                continue

            norm = token["norm"]
            if norm in self.lexicon:
                result.append(self.lexicon[norm])
            else:
                result.append(token["text"])

            i += 1

        return "".join(result)

    def _tokenize(self, text):
        parts = re.findall(r"[A-Za-z0-9_\-+/\.]+|[\s]+|[^\w\s]", text, flags=re.UNICODE)
        tokens = []

        for p in parts:
            if re.fullmatch(r"[A-Za-z0-9_\-+/\.]+", p):
                norm = p.lower()
                tokens.append({"type": "word", "text": p, "norm": norm})
            else:
                tokens.append({"type": "sep", "text": p, "norm": p})

        return tokens

    def _protect_special_tokens(self, text):
        patterns = [
            r"https?://[^\s]+",
            r"www\.[^\s]+",
            r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
            r"\bv?\d+(?:\.\d+){1,}\b",
        ]

        protected_map = {}
        protected_text = text

        counter = 0
        for pattern in patterns:
            matches = re.findall(pattern, protected_text)
            for m in matches:
                placeholder = f"__PROTECTED_{counter}__"
                protected_map[placeholder] = m
                protected_text = protected_text.replace(m, placeholder, 1)
                counter += 1

        return protected_map, protected_text

    def _restore_special_tokens(self, text, protected_map):
        result = text
        for placeholder, original in protected_map.items():
            result = result.replace(placeholder, original)
        return result

    def _cleanup_spaces(self, text):
        text = re.sub(r"\s+", " ", text)
        text = re.sub(r"\s+([,.;:!?])", r"\1", text)
        return text.strip()

    def _contains_chinese(self, text):
        return bool(re.search(r"[\u4e00-\u9fff]", text))
