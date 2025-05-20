import pandas as pd
import streamlit as st
import os
from datetime import datetime

def process_excel(input_file, output_file, selected_filters0, selected_filters1, selected_filters2, selected_filters3, group_by_columns):
    # 读取Excel文件，从第5行(index=4)开始加载数据
    df = pd.read_excel(input_file, header=0, skiprows=range(1, 4))
    
    # 记录原始数据以便后续恢复
    original_df = df.copy()
    
    # 记录处理前的总行数
    total_rows_before = len(df)
    
    # 获取列索引
    e_column = df.columns[4]  # 第5列
    f_column = df.columns[5]  # 第6列
    h_column = df.columns[7]  # 第8列 - 筛选条件1所在列
    i_column = df.columns[8]  # 第9列 - 筛选条件2所在列
    j_column = df.columns[9]  # 第10列 - 筛选条件3所在列
    k_column = df.columns[10]  # 第11列 - 筛选条件4所在列
    l_column = df.columns[11]  # 第12列
    
    # 保存原始索引顺序
    original_indices = df.index.tolist()
    
    # 应用筛选条件
    filtered_indices = set(df.index)
    
    if selected_filters0:
        filtered_indices &= set(df[df[h_column].isin(selected_filters0)].index)
    if selected_filters1:
        filtered_indices &= set(df[df[i_column].isin(selected_filters1)].index)
    if selected_filters2:
        filtered_indices &= set(df[df[j_column].isin(selected_filters2)].index)
    if selected_filters3:
        filtered_indices &= set(df[df[k_column].isin(selected_filters3)].index)
    
    # 按原始顺序筛选数据
    filtered_indices_ordered = [idx for idx in original_indices if idx in filtered_indices]
    filtered_df = df.loc[filtered_indices_ordered].copy()
    
    # 标记需要排除的行（值为0或"其他请注明"）
    exclude_mask = (filtered_df[k_column] == 0) | (filtered_df[k_column] == '其他请注明')
    
    # 映射列名到列索引
    column_mapping = {
        '筛选条件1 ': h_column,
        '筛选条件2': i_column,
        '筛选条件3（内容）': j_column,
        '筛选条件4': k_column
    }
    
    # 获取用户选择的分组列
    group_by_columns_indices = [column_mapping[col] for col in group_by_columns]
    
    # 初始化编号
    current_group_values = None
    current_number = 1
    numbers = []
    
    # 遍历筛选后的数据框，只处理不被排除的行
    for index, row in filtered_df.iterrows():
        if exclude_mask[index]:
            # 被排除的行，编号为None
            numbers.append(None)
            continue
        
        # 获取当前行的所有分组列的值
        group_values = tuple(row[col] for col in group_by_columns_indices)
        
        # 如果是第一行或分组列的值发生变化，增加编号
        if current_group_values is None:
            # 第一行的处理
            current_group_values = group_values
        elif group_values != current_group_values:
            # 分组列的值发生变化，增加编号
            current_number += 1
            current_group_values = group_values
        
        # 添加当前编号到列表
        numbers.append(current_number)
    
    # 确定Q列的位置
    q_column_index = 16  # 第17列
    
    # 确保DataFrame有足够的列
    while len(filtered_df.columns) <= q_column_index:
        filtered_df[f'Unnamed: {len(filtered_df.columns)}'] = None
    
    # 将编号添加到筛选后DataFrame的第17列(Q)
    filtered_df.iloc[:, q_column_index] = numbers
    
    # 根据第17列(Q)的编号进行分组处理，保留Q列为None的行
    groups = filtered_df.groupby(filtered_df.columns[q_column_index], dropna=False)
    
    # 处理12到16列的数据
    for group_number, group_df in groups:
        if group_number is None:
            # 处理被排除的行（Q为None）
            group_indices = group_df.index
            
            # 12列(L)设为空
            filtered_df.loc[group_indices, l_column] = ''
            
            # 13列(M)设为空
            m_column = filtered_df.columns[12]  # 第13列
            filtered_df.loc[group_indices, m_column] = ''
            
            # 14列(N)设为0
            n_column = filtered_df.columns[13]  # 第14列
            filtered_df.loc[group_indices, n_column] = 0
            
            # 15列(O)设为空
            o_column = filtered_df.columns[14]  # 第15列
            filtered_df.loc[group_indices, o_column] = ''
            
            # 16列(P)设为0
            p_column = filtered_df.columns[15]  # 第16列
            filtered_df.loc[group_indices, p_column] = '0'
            
            continue
        
        # 获取组内行索引
        group_indices = group_df.index
        
        # 生成26个英文字母(大写)
        letters = [chr(ord('A') + i) for i in range(26)]
        
        # 12列(L)：依次写入26个大写英文字母
        for i, idx in enumerate(group_indices):
            if i < len(letters):
                filtered_df.loc[idx, l_column] = letters[i]
            else:
                # 如果组内行数超过26，循环使用字母
                filtered_df.loc[idx, l_column] = letters[i % len(letters)]
        
        # 13列(M)：写入第10列对应行的值加上括号，括号里是12列字母
        m_column = filtered_df.columns[12]  # 第13列
        for idx in group_indices:
            letter = filtered_df.loc[idx, l_column]
            j_value = filtered_df.loc[idx, j_column]
            filtered_df.loc[idx, m_column] = f"{j_value}({letter})"
        
        # 14列(N)：写入第5列值的100倍，保留一位小数
        n_column = filtered_df.columns[13]  # 第14列
        for idx in group_indices:
            value = filtered_df.loc[idx, e_column]
            filtered_df.loc[idx, n_column] = round(value * 100, 1)
        
        # 15列(O)：写入第5列的值比这一行第6列的值小的同组其他行所对应的L列字母
        o_column = filtered_df.columns[14]  # 第15列
        for idx in group_indices:
            current_value = filtered_df.loc[idx, e_column]
            current_f_value = filtered_df.loc[idx, f_column]
            
            # 找出同组中第5列的值比当前行第6列的值小的行（排除自身）
            smaller_rows = group_df[(group_df[e_column] < current_f_value) & (group_df.index != idx)]
            
            # 收集这些行的L列字母
            letters_list = [filtered_df.loc[row_idx, l_column] for row_idx in smaller_rows.index]
            
            # 将字母列表合并为字符串
            filtered_df.loc[idx, o_column] = ''.join(letters_list)
        
        # 16列(P)：写入第14列和第15列合并的值，保留一位小数
        p_column = filtered_df.columns[15]  # 第16列
        for idx in group_indices:
            n_value = filtered_df.loc[idx, n_column]
            o_value = filtered_df.loc[idx, o_column]
            # 如果n_value是数字类型，保留一位小数
            if isinstance(n_value, (int, float)):
                n_value = round(n_value, 1)
            filtered_df.loc[idx, p_column] = f"{n_value}{o_value}"
    
    # 设置列名
    new_column_names = [
        '原题', 't', '自由度', '显著性 （双尾）', 
        '平均值差值', '95% 置信区间下限', '上限', '筛选条件1',
        '筛选条件2', '筛选条件3', '筛选条件4', 
        '字母', '内容+(字母)', '占比', 'sig win', '占比sig合并', '分组'
    ]
    
    # 确保列名数量与DataFrame列数匹配
    if len(new_column_names) == len(filtered_df.columns):
        filtered_df.columns = new_column_names
    else:
        st.warning(f"列名数量({len(new_column_names)})与实际列数({len(filtered_df.columns)})不匹配，将使用默认列名")
    
    # 保存处理后的筛选数据到新的Excel文件
    filtered_df.to_excel(output_file, index=False)
    
    # 构建分组依据的描述
    group_by_description = ", ".join(group_by_columns) if group_by_columns else "无（使用默认分组）"
    
    return {
        'total_rows_before': total_rows_before,
        'processed_rows': len(filtered_df) - sum(exclude_mask),
        'excluded_rows': sum(exclude_mask),
        'filtered_rows': total_rows_before - len(filtered_df),
        'group_by_description': group_by_description,
        'total_groups': current_number
    }

def main():
    st.title("Excel列处理工具")
    
    # 上传文件
    uploaded_file = st.file_uploader("请上传Excel文件", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # 保存上传的文件
        input_path = "input_file.xlsx"
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # 显示文件信息
        file_details = {"文件名": uploaded_file.name, "文件大小": uploaded_file.size}
        st.write(file_details)
        
        # 读取文件获取筛选条件
        df = pd.read_excel(input_path, header=0, skiprows=range(1, 4))
        
        # 获取筛选条件的唯一值
        filter0_options = df[df.columns[7]].dropna().unique().tolist()  # 第8列 - 范围列
        filter1_options = df[df.columns[8]].dropna().unique().tolist()
        filter2_options = df[df.columns[9]].dropna().unique().tolist()
        filter3_options = df[df.columns[10]].dropna().unique().tolist()
        
        # 创建筛选条件的多选框
        st.subheader("筛选条件")
        col1, col2 = st.columns(2)
        
        with col1:
            selected_filters0 = st.multiselect(
                "筛选条件1 ",
                filter0_options,
                default=[]
            )
        
        with col2:
            selected_filters1 = st.multiselect(
                "筛选条件2",
                filter1_options,
                default=[]
            )
        
        col3, col4 = st.columns(2)
        
        with col3:
            selected_filters2 = st.multiselect(
                "筛选条件3（内容）",
                filter2_options,
                default=[]
            )
        
        with col4:
            selected_filters3 = st.multiselect(
                "筛选条件4",
                filter3_options,
                default=[]
            )
        
        # 分组依据选择
        st.subheader("分组依据")
        group_by_columns = st.multiselect(
            "选择用于分组的列（选中列的值变化时创建新组）",
            [
                '筛选条件1 ',
                '筛选条件2',
                '筛选条件3（内容）',
                '筛选条件4'
            ],
            default=['筛选条件2', '筛选条件4']  # 默认使用原逻辑的分组依据
        )
        
        if not group_by_columns:
            st.warning("请至少选择一个分组依据列")
        
        # 处理按钮
        if st.button("开始处理") and group_by_columns:
            # 创建输出文件路径
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"output_file_{timestamp}.xlsx"
            
            # 处理文件
            with st.spinner("正在处理数据..."):
                result = process_excel(
                    input_path, output_path, 
                    selected_filters0, selected_filters1, 
                    selected_filters2, selected_filters3,
                    group_by_columns
                )
            
            # 显示处理结果
            st.success("处理完成!")
            st.write(f"总行数: {result['total_rows_before']}")
            st.write(f"筛选后剩余行数: {result['total_rows_before'] - result['filtered_rows']}")
            st.write(f"参与编号的行数: {result['processed_rows']}")
            st.write(f"被排除但保留的行数: {result['excluded_rows']}")
            st.write(f"分组依据: {result['group_by_description']}")
            st.write(f"共创建了 {result['total_groups']} 个分组")
            
            # 提供下载链接
            if os.path.exists(output_path):
                with open(output_path, "rb") as f:
                    bytes_data = f.read()
                st.download_button(
                    label="下载处理后的文件",
                    data=bytes_data,
                    file_name=os.path.basename(output_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # 清理临时文件
            os.remove(input_path)

if __name__ == "__main__":
    main()    
