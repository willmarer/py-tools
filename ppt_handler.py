import os
from pathlib import Path
from pptx import Presentation


class PPTProcessor:
    def __init__(self, translator, log_callback=None):
        self.translator = translator
        self.log = log_callback or (lambda msg: None)

    def translate_directory(self, input_dir, output_dir):
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        pptx_files = sorted(input_path.glob("*.pptx"))
        ppt_files = sorted(input_path.glob("*.ppt"))

        stats = {
            "total_files": len(pptx_files) + len(ppt_files),
            "success_files": 0,
            "failed_files": 0,
            "skipped_files": 0,
            "total_slides": 0,
            "translated_items": 0,
            "errors": [],
        }

        for f in ppt_files:
            self.log(f"[SKIP] {f.name} -> .ppt 暂不支持，请先另存为 .pptx")
            stats["skipped_files"] += 1

        for idx, file_path in enumerate(pptx_files, start=1):
            self.log("-" * 60)
            self.log(f"[FILE] ({idx}/{len(pptx_files)}) 开始处理: {file_path.name}")

            try:
                translated_count, slide_count, save_path = self.translate_pptx(file_path, output_path)
                stats["success_files"] += 1
                stats["total_slides"] += slide_count
                stats["translated_items"] += translated_count

                self.log(f"[DONE] 保存成功: {save_path}")
                self.log(f"[DONE] 页数: {slide_count}，翻译文本数: {translated_count}")

            except Exception as e:
                stats["failed_files"] += 1
                error_msg = f"{file_path.name} 处理失败: {e}"
                stats["errors"].append(error_msg)
                self.log(f"[ERROR] {error_msg}")

        return stats

    def translate_pptx(self, file_path, output_dir):
        prs = Presentation(str(file_path))
        slide_count = len(prs.slides)
        translated_count = 0

        for slide_index, slide in enumerate(prs.slides, start=1):
            self.log(f"  [SLIDE] {slide_index}/{slide_count}")
            translated_count += self._translate_slide(slide)

            if slide.has_notes_slide:
                try:
                    notes_frame = slide.notes_slide.notes_text_frame
                    translated_count += self._translate_text_frame(notes_frame)
                except Exception:
                    pass

        output_name = f"{file_path.stem}_translated.pptx"
        save_path = Path(output_dir) / output_name
        prs.save(str(save_path))

        return translated_count, slide_count, str(save_path)

    def _translate_slide(self, slide):
        translated_count = 0
        for shape in slide.shapes:
            translated_count += self._translate_shape(shape)
        return translated_count

    def _translate_shape(self, shape):
        translated_count = 0

        if hasattr(shape, "text_frame"):
            translated_count += self._translate_text_frame(shape.text_frame)

        if getattr(shape, "has_table", False):
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    translated_count += self._translate_text_frame(cell.text_frame)

        if getattr(shape, "shape_type", None) == 6 and hasattr(shape, "shapes"):
            for sub_shape in shape.shapes:
                translated_count += self._translate_shape(sub_shape)

        return translated_count

    def _translate_text_frame(self, text_frame):
        translated_count = 0

        if not text_frame:
            return translated_count

        for paragraph in text_frame.paragraphs:
            full_text = "".join(run.text for run in paragraph.runs)

            if not full_text or not full_text.strip():
                continue

            translated = self.translator.translate(full_text)

            if translated != full_text:
                paragraph.clear()
                paragraph.add_run().text = translated
                translated_count += 1

        return translated_count
