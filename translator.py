import json
import os
import re
import sys
import shutil


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_bundle_dir():
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

        self.normalized_phrases = self._build_normalized_phrases(self.phrases)
        self.max_phrase_len = self._calc_max_phrase_length()

    def _ensure_external_file(self, filename):
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

        normalized = {}
        for k, v in data.items():
            key = str(k).strip()
            value = str(v).strip()
            if key:
                normalized[key] = value
        return normalized

    def _build_normalized_phrases(self, phrases):
        result = {}
        for k, v in phrases.items():
            nk = self._normalize_phrase_key(k)
            if nk:
                result[nk] = v
        return result

    def _normalize_phrase_key(self, text):
        text = text.strip().lower()
        text = text.replace("-", " ")
        text = text.replace("_", " ")
        text = text.replace("/", " ")
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def _calc_max_phrase_length(self):
        max_len = 1
        for k in self.normalized_phrases.keys():
            word_count = len(k.split())
            if word_count > max_len:
                max_len = word_count
        return max_len

    def translate(self, text):
        if not text or not text.strip():
            return text

        original_text = text

        protected_map, protected_text = self._protect_special_tokens(text)
        translated = self._translate_core(protected_text)
        translated = self._restore_special_tokens(translated, protected_map)
        translated = self._cleanup_spaces(translated)

        if not translated.strip():
            return original_text

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

            # 1. 先做多 token 短语匹配（最长优先）
            for length in range(min(self.max_phrase_len, len(tokens) - i), 1, -1):
                phrase_match = self._try_match_phrase(tokens, i, length)
                if phrase_match is not None:
                    result.append(phrase_match["translation"])
                    i = phrase_match["next_index"]
                    matched = True
                    break

            if matched:
                continue

            norm = token["norm"]

            # 2. 再查普通单词表
            if norm in self.lexicon:
                result.append(self.lexicon[norm])
                i += 1
                continue

            # 3. 如果是带连接符的单 token，再尝试拆分后按短语/单词翻译
            split_translation = self._translate_compound_token(token["text"])
            if split_translation is not None:
                result.append(split_translation)
                i += 1
                continue

            # 4. 都没命中，保留原文
            result.append(token["text"])
            i += 1

        return "".join(result)

    def _translate_compound_token(self, token_text):
        """
        处理单个 token 内部带连接符的情况，例如：
        market_analysis
        market-analysis
        market/analysis
        """
        if not re.search(r"[_\-/]", token_text):
            return None

        normalized = self._normalize_phrase_key(token_text)
        if not normalized:
            return None

        # 先当短语查
        if normalized in self.normalized_phrases:
            return self.normalized_phrases[normalized]

        # 再拆成单词逐个查
        parts = normalized.split()
        translated_parts = []

        for p in parts:
            if p in self.lexicon:
                translated_parts.append(self.lexicon[p])
            else:
                translated_parts.append(p)

        return "".join(translated_parts)

    def _try_match_phrase(self, tokens, start_index, max_word_len):
        words = []
        j = start_index
        consumed_word_count = 0
        last_word_index = None

        while j < len(tokens) and consumed_word_count < max_word_len:
            t = tokens[j]

            if t["type"] == "word":
                words.append(t["norm"])
                consumed_word_count += 1
                last_word_index = j
                j += 1
                continue

            if t["type"] == "sep" and re.fullmatch(r"[\s\-_\/]+", t["text"]):
                j += 1
                continue

            break

        if consumed_word_count != max_word_len or last_word_index is None:
            return None

        phrase_key = " ".join(words).strip()
        phrase_key = self._normalize_phrase_key(phrase_key)

        if phrase_key in self.normalized_phrases:
            return {
                "translation": self.normalized_phrases[phrase_key],
                "next_index": last_word_index + 1
            }

        return None

    def _tokenize(self, text):
        pattern = r"__PROTECTED_\d+__|[A-Za-z0-9]+(?:[._+\-/][A-Za-z0-9]+)*|[\s]+|[^\w\s]"
        parts = re.findall(pattern, text, flags=re.UNICODE)

        tokens = []
        for p in parts:
            if re.fullmatch(r"__PROTECTED_\d+__", p):
                tokens.append({"type": "word", "text": p, "norm": p.lower()})
            elif re.fullmatch(r"[A-Za-z0-9]+(?:[._+\-/][A-Za-z0-9]+)*", p):
                tokens.append({"type": "word", "text": p, "norm": p.lower()})
            else:
                tokens.append({"type": "sep", "text": p, "norm": p})

        return tokens

    def _protect_special_tokens(self, text):
        patterns = [
            r"https?://[^\s]+",
            r"www\.[^\s]+",
            r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
            r"\bv?\d+(?:\.\d+){1,}\b",
            r"\b\d+(?:\.\d+)?%",
            r"\bQ[1-4]\b",
        ]

        protected_map = {}
        protected_text = text
        counter = 0

        for pattern in patterns:
            while True:
                match = re.search(pattern, protected_text)
                if not match:
                    break
                original = match.group(0)
                placeholder = f"__PROTECTED_{counter}__"
                protected_map[placeholder] = original
                protected_text = protected_text.replace(original, placeholder, 1)
                counter += 1

        return protected_map, protected_text

    def _restore_special_tokens(self, text, protected_map):
        result = text
        for placeholder, original in protected_map.items():
            result = result.replace(placeholder, original)
        return result

    def _cleanup_spaces(self, text):
        text = re.sub(r"[\t\r\f\v]+", " ", text)
        text = re.sub(r"\s+", " ", text)

        text = re.sub(r"([\u4e00-\u9fff])\s+([\u4e00-\u9fff])", r"\1\2", text)
        text = re.sub(r"([\u4e00-\u9fff])\s+([，。；：！？、】【（）《》“”‘’])", r"\1\2", text)
        text = re.sub(r"([，。；：！？、】【（）《》“”‘’])\s+([\u4e00-\u9fff])", r"\1\2", text)

        text = re.sub(r"\s+([,.;:!?])", r"\1", text)
        text = re.sub(r"([(\[{])\s+", r"\1", text)
        text = re.sub(r"\s+([)\]}])", r"\1", text)

        return text.strip()
