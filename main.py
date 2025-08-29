import pandas as pd
import numpy as np
import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import time


class ExamArrangementGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("考场安排系统")
        self.root.geometry("800x600")

        # 变量初始化
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.room_capacity = tk.IntVar(value=30)
        self.study_room_vars = {}
        self.random_seed = tk.IntVar(value=int(time.time()))  # 使用当前时间作为随机种子

        self.setup_ui()

    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 输入文件选择
        ttk.Label(main_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="浏览", command=self.select_input_file).grid(row=0, column=2)

        # 输出路径选择
        ttk.Label(main_frame, text="输出文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="浏览", command=self.select_output_folder).grid(row=1, column=2)

        # 考场容量设置
        ttk.Label(main_frame, text="每个考场人数:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(main_frame, from_=20, to=50, textvariable=self.room_capacity, width=10).grid(row=2, column=1,
                                                                                                 sticky=tk.W)

        # 随机种子设置
        ttk.Label(main_frame, text="随机种子:").grid(row=2, column=2, sticky=tk.W, pady=5, padx=5)
        ttk.Entry(main_frame, textvariable=self.random_seed, width=15).grid(row=2, column=3, sticky=tk.W)
        ttk.Button(main_frame, text="刷新种子", command=self.refresh_seed).grid(row=2, column=4, padx=5)

        # 自习室设置框架
        study_frame = ttk.LabelFrame(main_frame, text="自习室设置", padding="10")
        study_frame.grid(row=3, column=0, columnspan=5, sticky=(tk.W, tk.E), pady=10)

        # 选考科目列表
        elective_subjects = ['物理', '化学', '生物', '历史', '政治', '地理']

        # 创建复选框
        for i, subject in enumerate(elective_subjects):
            var = tk.BooleanVar(value=True)
            self.study_room_vars[subject] = var
            cb = ttk.Checkbutton(study_frame, text=f"{subject}科目安排自习室", variable=var)
            cb.grid(row=i // 3, column=i % 3, sticky=tk.W, padx=5, pady=2)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=5, pady=10)

        ttk.Button(button_frame, text="开始安排考场", command=self.run_arrangement).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT, padx=5)

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择学生名单文件",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.input_path.set(file_path)

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="选择输出文件夹")
        if folder_path:
            self.output_path.set(folder_path)

    def refresh_seed(self):
        """刷新随机种子"""
        self.random_seed.set(int(time.time()))

    def run_arrangement(self):
        if not self.input_path.get():
            messagebox.showerror("错误", "请选择输入文件")
            return

        if not self.output_path.get():
            messagebox.showerror("错误", "请选择输出文件夹")
            return

        try:
            # 获取用户设置的自习室安排
            need_study_room_dict = {}
            for subject, var in self.study_room_vars.items():
                need_study_room_dict[subject] = var.get()

            # 执行考场安排
            arrange_exam_rooms(
                self.input_path.get(),
                self.output_path.get(),
                self.room_capacity.get(),
                need_study_room_dict,
                self.random_seed.get()
            )

            messagebox.showinfo("完成", "考场安排已完成！")

        except Exception as e:
            messagebox.showerror("错误", f"发生错误: {str(e)}")


def arrange_exam_rooms(input_path, output_dir, room_capacity, need_study_room_dict, random_seed):
    """
    排考场主函数

    参数:
    input_path: 输入Excel文件路径
    output_dir: 输出文件夹路径
    room_capacity: 每个考场的考生数量
    need_study_room_dict: 字典，指定各科目是否需要为不参加考生安排自习室
    random_seed: 随机种子，用于确保结果可重现
    """

    # 设置随机种子
    np.random.seed(random_seed)
    random.seed(random_seed)

    # 读取Excel文件
    df = pd.read_excel(input_path, sheet_name=0)

    # 过滤掉备注为"不参加"的学生
    df = df[df['备注'] != '不参加']

    # 定义科目列表
    required_subjects = ['语文', '数学', '英语']  # 必考科目
    elective_subjects = ['物理', '化学', '生物', '历史', '政治', '地理']  # 选考科目
    all_subjects = required_subjects + elective_subjects

    # 为每个学生生成唯一的考场座位ID（用于必考科目）
    df['student_id'] = range(1, len(df) + 1)

    # 增强随机化：先对整个学生序列进行随机打乱
    df_shuffled = df.sample(frac=1, random_state=random_seed).reset_index(drop=True)

    # 初始化结果DataFrame - 使用原始数据框架，但后续会用打乱后的数据填充
    student_arrangements = df.copy()

    # 处理必考科目（语文、数学、英语）
    # 所有学生都必须参加必考科目，且使用相同的考场和座位
    required_students = df_shuffled.copy()
    required_students = assign_rooms_and_seats(required_students, room_capacity, "必考科目", random_seed)

    # 保存结果到学生安排表
    for subject in required_subjects:
        # 将打乱后的考场座位信息映射回原始顺序
        for idx, row in required_students.iterrows():
            original_idx = student_arrangements[student_arrangements['学籍号'] == row['学籍号']].index
            if len(original_idx) > 0:
                student_arrangements.loc[original_idx, f'{subject}考场'] = row['必考科目考场']
                student_arrangements.loc[original_idx, f'{subject}座位'] = row['必考科目座位']

    # 创建语数英教室安排表
    required_students_renamed = required_students.copy()
    required_students_renamed.rename(columns={
        '必考科目考场': '语数英考场',
        '必考科目座位': '语数英座位'
    }, inplace=True)

    room_arrangements = {}
    room_arrangements['语数英'] = create_room_arrangement_df(required_students_renamed, "语数英", False)

    # 处理选考科目
    for subject in elective_subjects:
        # 确定层次列名
        level_col = f'{subject}层次'

        # 筛选参加该科目考试的学生（层次不为空）
        subject_students = df_shuffled[df_shuffled[level_col].notna()].copy()

        # 按层次分组，尽量让同层次学生在同一考场
        subject_students = arrange_by_level(subject_students, level_col, room_capacity, subject, random_seed)

        # 保存结果到学生安排表
        for idx, row in subject_students.iterrows():
            original_idx = student_arrangements[student_arrangements['学籍号'] == row['学籍号']].index
            if len(original_idx) > 0:
                student_arrangements.loc[original_idx, f'{subject}考场'] = row[f'{subject}考场']
                student_arrangements.loc[original_idx, f'{subject}座位'] = row[f'{subject}座位']

        # 保存到教室安排表
        room_arrangements[subject] = create_room_arrangement_df(subject_students, subject, False)

        # 如果需要为不参加该科目的学生安排自习室
        if need_study_room_dict.get(subject, False):
            # 获取不参加该科目考试的学生
            non_subject_students = df_shuffled[df_shuffled[level_col].isna()].copy()

            if len(non_subject_students) > 0:
                # 获取当前科目的最大考场号
                max_room = subject_students[f'{subject}考场'].max() if not subject_students.empty else 0

                # 分配自习室
                non_subject_students = assign_study_rooms(non_subject_students, room_capacity, subject, max_room,
                                                          random_seed)

                # 保存到教室安排表（自习室）
                study_room_df = create_room_arrangement_df(non_subject_students, subject, True)

                # 合并到科目教室安排表
                room_arrangements[subject] = pd.concat([room_arrangements[subject], study_room_df], ignore_index=True)

                # 将自习室信息添加到学生安排表
                for idx, row in non_subject_students.iterrows():
                    original_idx = student_arrangements[student_arrangements['学籍号'] == row['学籍号']].index
                    if len(original_idx) > 0:
                        student_arrangements.loc[original_idx, f'{subject}考场'] = row[f'{subject}考场']
                        student_arrangements.loc[original_idx, f'{subject}座位'] = row[f'{subject}座位']

    # 创建输出文件夹
    output_dir = os.path.join(output_dir, "考场安排结果")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 保存学生安排表（按班级分sheet）
    student_output_path = os.path.join(output_dir, "学生考场安排表.xlsx")

    # 处理考场座位格式
    # 合并语文、数学、英语三科的考场座位信息
    student_arrangements['语数英考场座位号'] = student_arrangements.apply(
        lambda row: f"{int(row['语文考场'])}考场{int(row['语文座位'])}座"
        if pd.notna(row['语文考场']) and pd.notna(row['语文座位']) else "", axis=1
    )

    # 为其他科目创建考场座位合并列
    for subject in elective_subjects:
        # 先创建一个临时列，用于判断是否有考场或自习室安排
        student_arrangements[f'{subject}安排'] = student_arrangements.apply(
            lambda row, subj=subject:
            f"{int(row[f'{subj}考场'])}考场{int(row[f'{subj}座位'])}座"
            if pd.notna(row[f'{subj}考场']) and pd.notna(row[f'{subj}座位']) else
            f"{int(row[f'{subj}考场'])}考场自习"  # 修正为"X考场自习"格式
            if pd.notna(row[f'{subj}考场']) and pd.isna(row[f'{subj}座位']) else
            "", axis=1
        )

    # 选择需要输出的列
    output_columns = ['校区', '年级', '班级', '姓名', '学籍号', '语数英考场座位号']
    for subject in elective_subjects:
        output_columns.append(f'{subject}安排')

    # 创建列名映射（简化列名）
    column_mapping = {
        '语数英考场座位号': '语数英'
    }
    for subject in elective_subjects:
        column_mapping[f'{subject}安排'] = subject

    # 创建输出DataFrame（使用简化列名）
    output_df = student_arrangements[output_columns].copy()
    output_df.rename(columns=column_mapping, inplace=True)

    # 按班级分组并保存到不同sheet
    with pd.ExcelWriter(student_output_path) as writer:
        # 先保存所有学生的总表
        output_df.to_excel(writer, sheet_name='所有学生', index=False)

        # 按班级分组保存
        classes = student_arrangements['班级'].unique()
        for class_name in classes:
            class_df = output_df[student_arrangements['班级'] == class_name]
            # 处理sheet名称，避免过长或非法字符
            sheet_name = f'班级{class_name}'
            if len(sheet_name) > 31:  # Excel sheet名称长度限制
                sheet_name = sheet_name[:31]
            class_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # 保存教室安排表（按教室分sheet）
    room_output_path = os.path.join(output_dir, "教室安排表.xlsx")
    with pd.ExcelWriter(room_output_path) as writer:
        # 处理语数英科目
        if '语数英' in room_arrangements:
            df_room = room_arrangements['语数英']
            # 按教室编号分组
            for room_num, group in df_room.groupby('教室编号'):
                sheet_name = f"语数英_教室{room_num}"
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                group.to_excel(writer, sheet_name=sheet_name, index=False)

        # 处理其他科目
        for subject in elective_subjects:
            if subject in room_arrangements:
                df_room = room_arrangements[subject]
                # 按教室编号和类型分组
                for (room_num, room_type), group in df_room.groupby(['教室编号', '类型']):
                    sheet_name = f"{subject}_{room_type}_教室{room_num}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    group.to_excel(writer, sheet_name=sheet_name, index=False)

    # 保存随机种子信息
    seed_info_path = os.path.join(output_dir, "随机种子信息.txt")
    with open(seed_info_path, 'w') as f:
        f.write(f"本次考场安排使用的随机种子: {random_seed}\n")
        f.write(f"使用相同的种子可以重现相同的考场安排结果\n")


def assign_rooms_and_seats(students_df, room_capacity, subject, random_seed):
    """为指定科目的学生分配考场和座位"""
    # 随机打乱学生顺序（增加随机性）
    students_df = students_df.sample(frac=1, random_state=random_seed).reset_index(drop=True)

    # 计算需要的考场数量
    num_students = len(students_df)
    num_rooms = (num_students + room_capacity - 1) // room_capacity

    # 分配考场和座位
    rooms = []
    seats = []

    for i in range(num_students):
        room_num = (i // room_capacity) + 1
        seat_num = (i % room_capacity) + 1
        rooms.append(room_num)
        seats.append(seat_num)

    students_df[f'{subject}考场'] = rooms
    students_df[f'{subject}座位'] = seats

    return students_df


def arrange_by_level(students_df, level_col, room_capacity, subject, random_seed):
    """按层次分组安排考场，尽量让同层次学生在同一考场"""
    # 按层次分组
    grouped = students_df.groupby(level_col)

    result_dfs = []

    for level, group in grouped:
        # 打乱同一层次内的学生顺序（增加随机性）
        group_seed = random_seed + hash(level) % 1000  # 为不同层次使用不同的随机种子
        group = group.sample(frac=1, random_state=group_seed).reset_index(drop=True)

        # 分配考场和座位
        num_students = len(group)
        rooms = []
        seats = []

        for i in range(num_students):
            room_num = (i // room_capacity) + 1
            seat_num = (i % room_capacity) + 1
            rooms.append(room_num)
            seats.append(seat_num)

        group[f'{subject}考场'] = rooms
        group[f'{subject}座位'] = seats

        result_dfs.append(group)

    # 合并所有层次的结果
    result_df = pd.concat(result_dfs, ignore_index=True)

    # 重新调整考场号，确保连续
    room_mapping = {}
    current_room = 1
    for room in sorted(result_df[f'{subject}考场'].unique()):
        room_mapping[room] = current_room
        current_room += 1

    result_df[f'{subject}考场'] = result_df[f'{subject}考场'].map(room_mapping)

    return result_df


def assign_study_rooms(students_df, room_capacity, subject, start_room, random_seed):
    """为不参加考试的学生分配自习室"""
    # 打乱学生顺序（增加随机性）
    students_df = students_df.sample(frac=1, random_state=random_seed).reset_index(drop=True)

    # 计算需要的自习室数量
    num_students = len(students_df)
    num_rooms = (num_students + room_capacity - 1) // room_capacity

    # 分配自习室（不分配具体座位）
    rooms = []

    for i in range(num_students):
        room_num = start_room + (i // room_capacity) + 1
        rooms.append(room_num)

    students_df[f'{subject}考场'] = rooms
    students_df[f'{subject}座位'] = None  # 自习室不分配固定座位

    return students_df


def create_room_arrangement_df(students_df, subject, is_study_room):
    """创建教室安排DataFrame"""
    # 根据科目名称确定列名
    if subject == "语数英":
        room_col = "语数英考场"
        seat_col = "语数英座位"
    else:
        room_col = f'{subject}考场'
        seat_col = f'{subject}座位'

    if is_study_room:
        room_type = "自习室"
        # 自习室只需要教室、姓名和学籍号
        result_df = students_df[['姓名', '学籍号', room_col]].copy()
        result_df['座位号'] = None
    else:
        room_type = "考场"
        # 考场需要教室、姓名、学籍号和座位号
        result_df = students_df[['姓名', '学籍号', room_col, seat_col]].copy()
        result_df.rename(columns={seat_col: '座位号'}, inplace=True)

    result_df['科目'] = subject
    result_df['类型'] = room_type
    result_df.rename(columns={room_col: '教室编号'}, inplace=True)

    # 重新排列列顺序
    result_df = result_df[['科目', '类型', '教室编号', '姓名', '学籍号', '座位号']]

    # 按教室编号和座位号排序
    if not is_study_room:
        result_df = result_df.sort_values(['教室编号', '座位号'])
    else:
        result_df = result_df.sort_values(['教室编号'])

    return result_df


if __name__ == "__main__":
    app = ExamArrangementGUI()
    app.root.mainloop()