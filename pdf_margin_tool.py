import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz  # PyMuPDF
import os
import sys
from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFMarginTool:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 여백 & 테두리 추가 + N쪽 모아찍기")
        self.root.geometry("550x830")
        
        # 윈도우 아이콘 설정
        self.setup_icon()
        
        self.pdf_path = None
        
        # 제목
        title = tk.Label(root, text="PDF 여백 & 테두리 + N쪽 모아찍기", 
                        font=("맑은 고딕", 16, "bold"))
        title.pack(pady=15)
        
        # 드래그앤드롭 영역
        drop_frame = tk.Frame(root, bg="#e8f4f8", relief="solid", borderwidth=2)
        drop_frame.pack(pady=10, padx=20, fill="x")
        
        drop_label = tk.Label(drop_frame, 
                             text="📄 PDF 파일을 여기에 드래그하세요",
                             font=("맑은 고딕", 12),
                             bg="#e8f4f8",
                             fg="#2196F3",
                             pady=15)
        drop_label.pack()
        
        # 드래그앤드롭 등록
        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', self.drop_file)
        drop_label.drop_target_register(DND_FILES)
        drop_label.dnd_bind('<<Drop>>', self.drop_file)
        
        # 또는 라벨
        tk.Label(root, text="또는", font=("맑은 고딕", 9), fg="gray").pack(pady=5)
        
        # PDF 선택 버튼
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)
        
        self.select_btn = tk.Button(btn_frame, text="PDF 파일 선택", 
                                    command=self.select_pdf,
                                    width=25, height=1,
                                    bg="#4CAF50", fg="white",
                                    font=("맑은 고딕", 12, "bold"))
        self.select_btn.pack()
        
        # 선택된 파일 표시
        self.file_label = tk.Label(root, text="선택된 파일 없음", 
                                   fg="gray", wraplength=500,
                                   font=("맑은 고딕", 10))
        self.file_label.pack(pady=5)
        
        # 설정 프레임
        settings_frame = tk.LabelFrame(root, text="설정", 
                                      font=("맑은 고딕", 10, "bold"),
                                      padx=20, pady=15)
        settings_frame.pack(pady=10, padx=20, fill="both")
        
        # 안쪽 여백 (내용과 테두리 사이)
        inner_frame = tk.Frame(settings_frame)
        inner_frame.pack(pady=5)
        
        tk.Label(inner_frame, text="안쪽 여백 (mm):", 
                font=("맑은 고딕", 9), width=15, anchor="w").pack(side="left")
        self.inner_margin_var = tk.StringVar(value="2")
        inner_entry = tk.Entry(inner_frame, textvariable=self.inner_margin_var, 
                              width=10)
        inner_entry.pack(side="left", padx=5)
        tk.Label(inner_frame, text="(내용↔테두리)", 
                fg="gray", font=("맑은 고딕", 8)).pack(side="left")
        
        # 바깥쪽 여백 (테두리 바깥)
        outer_frame = tk.Frame(settings_frame)
        outer_frame.pack(pady=5)
        
        tk.Label(outer_frame, text="바깥쪽 여백 (mm):", 
                font=("맑은 고딕", 9), width=15, anchor="w").pack(side="left")
        self.outer_margin_var = tk.StringVar(value="8")
        outer_entry = tk.Entry(outer_frame, textvariable=self.outer_margin_var, 
                              width=10)
        outer_entry.pack(side="left", padx=5)
        tk.Label(outer_frame, text="(테두리 바깥)", 
                fg="gray", font=("맑은 고딕", 8)).pack(side="left")
        
        # 테두리 설정
        border_frame = tk.Frame(settings_frame)
        border_frame.pack(pady=5)
        
        tk.Label(border_frame, text="테두리 두께:", 
                font=("맑은 고딕", 9), width=15, anchor="w").pack(side="left")
        self.border_var = tk.StringVar(value="1")
        border_entry = tk.Entry(border_frame, textvariable=self.border_var, 
                               width=10)
        border_entry.pack(side="left", padx=5)
        
        # N쪽 모아찍기 옵션
        nup_frame = tk.LabelFrame(settings_frame, text="모아찍기 옵션",
                                 font=("맑은 고딕", 9, "bold"),
                                 padx=10, pady=10)
        nup_frame.pack(pady=10, fill="x")
        
        self.nup_enable_var = tk.BooleanVar(value=True)
        nup_check = tk.Checkbutton(nup_frame, 
                                   text="모아찍기 적용", 
                                   variable=self.nup_enable_var,
                                   font=("맑은 고딕", 10, "bold"),
                                   command=self.toggle_nup_options)
        nup_check.pack(anchor="w", pady=5)
        
        # 용지 크기 선택
        paper_frame = tk.Frame(nup_frame)
        paper_frame.pack(pady=5, fill="x")
        
        tk.Label(paper_frame, text="용지 크기:", 
                font=("맑은 고딕", 9)).pack(side="left", padx=(0, 10))
        
        self.paper_var = tk.StringVar(value="A4")
        papers = [("A4", "A4"), ("A3", "A3"), ("Letter", "Letter"), ("자동", "auto")]
        for text, value in papers:
            rb = tk.Radiobutton(paper_frame, 
                              text=text, 
                              variable=self.paper_var,
                              value=value,
                              font=("맑은 고딕", 9))
            rb.pack(side="left", padx=5)
        
        self.paper_frame = paper_frame
        
        # 종이 방향
        orientation_frame = tk.Frame(nup_frame)
        orientation_frame.pack(pady=5, fill="x")
        
        tk.Label(orientation_frame, text="종이 방향:", 
                font=("맑은 고딕", 9)).pack(side="left", padx=(0, 10))
        
        self.orientation_var = tk.StringVar(value="portrait")
        tk.Radiobutton(orientation_frame, 
                      text="세로", 
                      variable=self.orientation_var,
                      value="portrait",
                      font=("맑은 고딕", 9)).pack(side="left", padx=5)
        tk.Radiobutton(orientation_frame, 
                      text="가로", 
                      variable=self.orientation_var,
                      value="landscape",
                      font=("맑은 고딕", 9)).pack(side="left", padx=5)
        
        self.orientation_frame = orientation_frame
        
        # 페이지 간격
        spacing_frame = tk.Frame(nup_frame)
        spacing_frame.pack(pady=5, fill="x")
        
        tk.Label(spacing_frame, text="페이지 간격 (mm):", 
                font=("맑은 고딕", 9)).pack(side="left", padx=(0, 10))
        
        self.spacing_var = tk.StringVar(value="5")
        spacing_entry = tk.Entry(spacing_frame, textvariable=self.spacing_var, 
                                width=10)
        spacing_entry.pack(side="left", padx=5)
        tk.Label(spacing_frame, text="(좌우상하)", 
                fg="gray", font=("맑은 고딕", 8)).pack(side="left")
        
        self.spacing_frame = spacing_frame
        
        # 레이아웃 설정 (행, 열)
        layout_frame = tk.Frame(nup_frame)
        layout_frame.pack(pady=8, fill="x")
        
        tk.Label(layout_frame, text="레이아웃:", 
                font=("맑은 고딕", 9)).pack(side="left", padx=(0, 10))
        
        tk.Label(layout_frame, text="행:", 
                font=("맑은 고딕", 9)).pack(side="left", padx=5)
        self.rows_var = tk.StringVar(value="2")
        rows_entry = tk.Entry(layout_frame, textvariable=self.rows_var, 
                             width=5)
        rows_entry.pack(side="left", padx=5)
        
        tk.Label(layout_frame, text="열:", 
                font=("맑은 고딕", 9)).pack(side="left", padx=5)
        self.cols_var = tk.StringVar(value="1")
        cols_entry = tk.Entry(layout_frame, textvariable=self.cols_var, 
                             width=5)
        cols_entry.pack(side="left", padx=5)
        
        tk.Label(layout_frame, text="(예: 2x1 = 상하)", 
                fg="gray", font=("맑은 고딕", 8)).pack(side="left", padx=5)
        
        self.layout_frame = layout_frame
        
        # 프리셋 버튼들
        preset_frame = tk.Frame(nup_frame)
        preset_frame.pack(pady=5, fill="x")
        
        tk.Label(preset_frame, text="프리셋:", 
                font=("맑은 고딕", 9)).pack(side="left", padx=(0, 10))
        
        presets = [
            ("2쪽", "2", "1", "portrait"),
            ("4쪽", "2", "2", "landscape"),
            ("6쪽", "2", "3", "landscape"),
            ("8쪽", "4", "2", "portrait"),
        ]
        
        for name, rows, cols, orient in presets:
            btn = tk.Button(preset_frame, text=name,
                          command=lambda r=rows, c=cols, o=orient: self.set_preset(r, c, o),
                          font=("맑은 고딕", 8),
                          width=6)
            btn.pack(side="left", padx=2)
        
        self.preset_frame = preset_frame
        
        # 처리 버튼
        self.process_btn = tk.Button(root, text="PDF 처리 및 저장", 
                                     command=self.process_pdf,
                                     width=25, height=2,
                                     bg="#2196F3", fg="white",
                                     font=("맑은 고딕", 12, "bold"),
                                     state="disabled")
        self.process_btn.pack(pady=10)
        
        # 진행 상태
        self.status_label = tk.Label(root, text="", fg="green", 
                                    font=("맑은 고딕", 10, "bold"))
        self.status_label.pack(pady=3)
        
        # 하단 제작자 정보
        footer_frame = tk.Frame(root, bg="#f0f0f0", height=30)
        footer_frame.pack(side="bottom", fill="x", pady=(10, 0))
        
        footer_label = tk.Label(footer_frame, 
                               text="제작: 휴먼임팩트협동조합", 
                               font=("맑은 고딕", 9),
                               bg="#f0f0f0",
                               fg="#666666")
        footer_label.pack(pady=5)
    
    def setup_icon(self):
        """아이콘 설정"""
        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        icon_loaded = False
        
        png_path = os.path.join(application_path, "icon.png")
        if not icon_loaded and os.path.exists(png_path):
            try:
                icon = tk.PhotoImage(file=png_path)
                self.root.iconphoto(True, icon)
                self.root.icon = icon
                icon_loaded = True
            except:
                pass
        
        ico_path = os.path.join(application_path, "icon.ico")
        if not icon_loaded and os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
            except:
                pass
        
        if not icon_loaded:
            for icon_file in ["icon.png", "icon.ico"]:
                if os.path.exists(icon_file):
                    try:
                        if icon_file.endswith('.png'):
                            icon = tk.PhotoImage(file=icon_file)
                            self.root.iconphoto(True, icon)
                            self.root.icon = icon
                        else:
                            self.root.iconbitmap(icon_file)
                        icon_loaded = True
                        break
                    except:
                        pass
    
    def drop_file(self, event):
        """드래그앤드롭으로 파일 선택"""
        file_path = event.data
        
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        if file_path.lower().endswith('.pdf'):
            self.pdf_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"✓ 선택: {filename}", fg="green", font=("맑은 고딕", 9, "bold"))
            self.process_btn.config(state="normal")
        else:
            messagebox.showerror("오류", "PDF 파일만 선택할 수 있습니다.")
    
    def set_preset(self, rows, cols, orientation):
        """프리셋 적용"""
        self.rows_var.set(rows)
        self.cols_var.set(cols)
        self.orientation_var.set(orientation)
    
    def toggle_nup_options(self):
        """모아찍기 옵션 토글"""
        state = "normal" if self.nup_enable_var.get() else "disabled"
        
        for child in self.paper_frame.winfo_children():
            if isinstance(child, tk.Radiobutton):
                child.config(state=state)
        
        for child in self.orientation_frame.winfo_children():
            if isinstance(child, tk.Radiobutton):
                child.config(state=state)
        
        for child in self.spacing_frame.winfo_children():
            if isinstance(child, tk.Entry):
                child.config(state=state)
        
        for child in self.layout_frame.winfo_children():
            if isinstance(child, tk.Entry):
                child.config(state=state)
        
        for child in self.preset_frame.winfo_children():
            if isinstance(child, tk.Button):
                child.config(state=state)
    
    def select_pdf(self):
        """PDF 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file_path:
            self.pdf_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"✓ 선택: {filename}", fg="green", font=("맑은 고딕", 9, "bold"))
            self.process_btn.config(state="normal")
    
    def get_paper_size(self, paper_format, orientation):
        """용지 크기 반환"""
        sizes = {
            "A4": (595, 842),
            "A3": (842, 1191),
            "Letter": (612, 792),
        }
        
        if paper_format == "auto":
            return None
        
        width, height = sizes.get(paper_format, (595, 842))
        
        if orientation == "landscape":
            return (height, width)
        return (width, height)
    
    def copy_links(self, src_page, dst_page, src_rect, dst_rect):
        """링크 복사"""
        try:
            links = src_page.get_links()
            
            scale_x = dst_rect.width / src_rect.width if src_rect.width > 0 else 1
            scale_y = dst_rect.height / src_rect.height if src_rect.height > 0 else 1
            
            offset_x = dst_rect.x0 - src_rect.x0 * scale_x
            offset_y = dst_rect.y0 - src_rect.y0 * scale_y
            
            for link in links:
                try:
                    link_rect = link.get("from", None)
                    if not link_rect:
                        continue
                    
                    new_rect = fitz.Rect(
                        link_rect.x0 * scale_x + offset_x,
                        link_rect.y0 * scale_y + offset_y,
                        link_rect.x1 * scale_x + offset_x,
                        link_rect.y1 * scale_y + offset_y
                    )
                    
                    new_link = {
                        "kind": link.get("kind", 2),
                        "from": new_rect,
                    }
                    
                    if "uri" in link:
                        new_link["uri"] = link["uri"]
                    if "page" in link:
                        new_link["page"] = link["page"]
                    if "to" in link:
                        new_link["to"] = link["to"]
                    if "file" in link:
                        new_link["file"] = link["file"]
                    if "zoom" in link:
                        new_link["zoom"] = link["zoom"]
                    
                    dst_page.insert_link(new_link)
                except:
                    continue
        except:
            pass
    
    def process_pdf(self):
        """PDF 처리"""
        if not self.pdf_path:
            messagebox.showerror("오류", "PDF 파일을 먼저 선택하세요.")
            return
        
        try:
            inner_mm = float(self.inner_margin_var.get())
            outer_mm = float(self.outer_margin_var.get())
            border_width = int(self.border_var.get())
            do_nup = self.nup_enable_var.get()
            rows = int(self.rows_var.get())
            cols = int(self.cols_var.get())
            orientation = self.orientation_var.get()
            paper_format = self.paper_var.get()
            spacing_mm = float(self.spacing_var.get())
            
            if inner_mm < 0 or outer_mm < 0 or border_width < 0 or spacing_mm < 0:
                messagebox.showerror("오류", "양수 값을 입력하세요.")
                return
            
            if do_nup and (rows <= 0 or cols <= 0):
                messagebox.showerror("오류", "행과 열은 1 이상이어야 합니다.")
                return
        except ValueError:
            messagebox.showerror("오류", "올바른 숫자를 입력하세요.")
            return
        
        # 원본 파일 경로와 이름
        original_dir = os.path.dirname(self.pdf_path)
        original_name = os.path.basename(self.pdf_path)
        name_without_ext = os.path.splitext(original_name)[0]
        
        # 기본 파일명
        nup_count = rows * cols
        if do_nup:
            default_name = f"{name_without_ext}_processed_{nup_count}up.pdf"
        else:
            default_name = f"{name_without_ext}_margin.pdf"
        
        # 저장 경로
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=default_name,
            initialdir=original_dir
        )
        
        if not save_path:
            return
        
        try:
            self.status_label.config(text="처리 중...", fg="orange")
            self.root.update()
            
            inner = inner_mm * 2.83465
            outer = outer_mm * 2.83465
            spacing = spacing_mm * 2.83465
            
            doc = fitz.open(self.pdf_path)
            margin_doc = fitz.open()
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                old_rect = page.rect
                
                new_width = old_rect.width + 2 * (inner + outer)
                new_height = old_rect.height + 2 * (inner + outer)
                
                new_page = margin_doc.new_page(width=new_width, height=new_height)
                
                content_rect = fitz.Rect(
                    outer + inner, 
                    outer + inner, 
                    new_width - outer - inner, 
                    new_height - outer - inner
                )
                
                new_page.show_pdf_page(content_rect, doc, page_num)
                self.copy_links(page, new_page, old_rect, content_rect)
                
                if border_width > 0:
                    shape = new_page.new_shape()
                    shape.draw_rect(fitz.Rect(
                        outer,
                        outer,
                        new_width - outer,
                        new_height - outer
                    ))
                    shape.finish(width=border_width, color=(0, 0, 0))
                    shape.commit()
            
            doc.close()
            
            if do_nup:
                self.status_label.config(text=f"{nup_count}쪽 모아찍기 중...", fg="orange")
                self.root.update()
                
                final_doc = fitz.open()
                total_pages = len(margin_doc)
                sample_page = margin_doc[0]
                page_width = sample_page.rect.width
                page_height = sample_page.rect.height
                paper_size = self.get_paper_size(paper_format, orientation)
                
                if paper_size:
                    paper_width, paper_height = paper_size
                    total_spacing_width = spacing * (cols - 1)
                    total_spacing_height = spacing * (rows - 1)
                    available_width = paper_width - total_spacing_width
                    available_height = paper_height - total_spacing_height
                    scale_width = available_width / (page_width * cols)
                    scale_height = available_height / (page_height * rows)
                    scale = min(scale_width, scale_height)
                    scaled_page_width = page_width * scale
                    scaled_page_height = page_height * scale
                    cell_width = paper_width / cols
                    cell_height = paper_height / rows
                else:
                    paper_width = page_width * cols + spacing * (cols - 1)
                    paper_height = page_height * rows + spacing * (rows - 1)
                    scaled_page_width = page_width
                    scaled_page_height = page_height
                    cell_width = page_width + spacing / cols
                    cell_height = page_height + spacing / rows
                
                for group_start in range(0, total_pages, nup_count):
                    combined_page = final_doc.new_page(width=paper_width, height=paper_height)
                    
                    for i in range(nup_count):
                        page_idx = group_start + i
                        if page_idx >= total_pages:
                            break
                        
                        row = i // cols
                        col = i % cols
                        cell_x = col * cell_width
                        cell_y = row * cell_height
                        x_offset = (cell_width - scaled_page_width) / 2
                        y_offset = (cell_height - scaled_page_height) / 2
                        x = cell_x + x_offset
                        y = cell_y + y_offset
                        dst_rect = fitz.Rect(x, y, x + scaled_page_width, y + scaled_page_height)
                        src_page = margin_doc[page_idx]
                        combined_page.show_pdf_page(dst_rect, margin_doc, page_idx)
                        self.copy_links(src_page, combined_page, src_page.rect, dst_rect)
                
                margin_doc.close()
                final_doc.save(save_path)
                final_doc.close()
            else:
                margin_doc.save(save_path)
                margin_doc.close()
            
            self.status_label.config(text="완료! ✓", fg="green")
            
            if do_nup:
                paper_info = f"{paper_format} " if paper_format != "auto" else ""
                messagebox.showinfo("완료", f"파일이 저장되었습니다!\n\n여백, 테두리, {rows}x{cols} 모아찍기가\n{paper_info}용지에 적용되었습니다.")
            else:
                messagebox.showinfo("완료", f"파일이 저장되었습니다!\n\n여백과 테두리가 적용되었습니다.")
        except Exception as e:
            self.status_label.config(text="오류 발생", fg="red")
            messagebox.showerror("오류", f"처리 중 오류 발생:\n{str(e)}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFMarginTool(root)
    root.mainloop()
