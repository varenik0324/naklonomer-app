    @classmethod
    def parse_settlement(cls, file_bytes: bytes, sheet_name: str,
                         corner_marks: List[str], L_fund: float, B_fund: float,
                         mark_col: int = 0) -> Tuple[Optional[pd.DataFrame], List[str]]:
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        except Exception as e:
            st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
            return None, []

        header_row, cycle_names = cls._extract_cycle_headers(df_raw)
        if header_row is None:
            st.error("Не найдена строка с заголовками циклов в листе осадок.")
            return None, []
        cycle_labels = cls._parse_cycle_labels(cycle_names)
        if not cycle_labels:
            st.error("Не удалось распознать даты циклов в осадках.")
            return None, []

        # Находим строки с марками (числа в столбце mark_col)
        mark_rows = []
        for idx in range(header_row + 1, len(df_raw)):
            val = df_raw.iloc[idx, mark_col]
            if pd.notna(val) and isinstance(val, (int, float)):
                mark_rows.append(idx)

        if not mark_rows:
            st.error(f"Не найдены строки с марками в столбце {mark_col}.")
            return None, []

        # Собираем все доступные марки
        available_marks = []
        for idx in mark_rows:
            mark_num = df_raw.iloc[idx, mark_col]
            if isinstance(mark_num, (int, float)):
                mark_str = str(int(mark_num)) if mark_num == int(mark_num) else str(mark_num)
                available_marks.append(mark_str)

        # Для каждого цикла определяем столбец с осадками:
        # ищем столбец справа от заголовка цикла, который содержит числа в строках марок
        cycle_to_col = {}
        total_cols = len(df_raw.columns)
        for label in cycle_labels:
            # Найдём столбец, где в заголовке есть этот цикл
            col_idx = None
            for c in range(total_cols):
                cell = df_raw.iloc[header_row, c]
                if pd.notna(cell) and label in str(cell):
                    col_idx = c
                    break
            if col_idx is None:
                st.warning(f"Не найден заголовок для цикла {label}")
                continue

            # Ищем столбец с осадками: начиная с col_idx+1, ищем столбец, где есть числа в строках марок
            found_col = None
            for c in range(col_idx + 1, min(col_idx + 5, total_cols)):  # проверяем следующие 4 столбца
                # Проверяем, есть ли числа в этом столбце для марок
                has_numbers = False
                for r in mark_rows:
                    val = df_raw.iloc[r, c]
                    if pd.notna(val) and isinstance(val, (int, float)):
                        has_numbers = True
                        break
                if has_numbers:
                    found_col = c
                    break

            if found_col is not None:
                cycle_to_col[label] = found_col
            else:
                st.warning(f"Не найден столбец с осадками для цикла {label}")

        if not cycle_to_col:
            st.error("Не найдены столбцы с осадками. Убедитесь, что в листе 'Стилобат' после каждого заголовка цикла идут числовые значения осадок.")
            return None, available_marks

        # Собираем данные осадок
        marks_data = {}
        for cycle, col_idx in cycle_to_col.items():
            marks_data[cycle] = {}
            for row_idx in mark_rows:
                mark_num = df_raw.iloc[row_idx, mark_col]
                mark_str = str(int(mark_num)) if mark_num == int(mark_num) else str(mark_num)
                sett_val = df_raw.iloc[row_idx, col_idx]
                if pd.notna(sett_val) and isinstance(sett_val, (int, float)):
                    marks_data[cycle][mark_str] = sett_val
                else:
                    marks_data[cycle][mark_str] = np.nan

        sorted_cycles = sorted(cycle_labels)
        zero_cycle = sorted_cycles[0] if sorted_cycles else None
        if zero_cycle is None:
            st.error("Нет циклов для определения нулевого.")
            return None, available_marks

        corner_marks_str = [str(m) for m in corner_marks]
        for mark in corner_marks_str:
            if mark not in marks_data[zero_cycle] or np.isnan(marks_data[zero_cycle][mark]):
                st.warning(f"Марка {mark} отсутствует в нулевом цикле {zero_cycle}. Попробуйте другие марки.")
                return None, available_marks

        results = []
        for cycle in sorted_cycles:
            if cycle == zero_cycle:
                continue
            s = {}
            for mark in corner_marks_str:
                if mark in marks_data[cycle] and mark in marks_data[zero_cycle]:
                    s[mark] = marks_data[cycle][mark] - marks_data[zero_cycle][mark]
                else:
                    s[mark] = np.nan
            if any(np.isnan(list(s.values()))):
                continue
            s1, s2, s3, s4 = s[corner_marks_str[0]], s[corner_marks_str[1]], s[corner_marks_str[2]], s[corner_marks_str[3]]
            a = ((s2 - s1) + (s4 - s3)) / (2 * L_fund) if L_fund != 0 else 0
            b = ((s3 - s1) + (s4 - s2)) / (2 * B_fund) if B_fund != 0 else 0
            results.append({
                'Цикл': cycle,
                'a_мм_м': a,
                'b_мм_м': b,
                'a_град': np.degrees(np.arctan(a / 1000)),
                'b_град': np.degrees(np.arctan(b / 1000))
            })

        if not results:
            st.error("Не удалось рассчитать углы по осадкам.")
            return None, available_marks

        return pd.DataFrame(results), available_marks
