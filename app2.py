with tab2:
    st.subheader("📈 Изменение абсолютных углов наклона по циклам")

    # Получаем выбранные этажи из session_state (установлены в tab1)
    selected_floors = st.session_state.get('floor_selector', sorted(df_incl['Этаж'].unique()))
    df_filtered = df_incl[df_incl['Этаж'].isin(selected_floors)]

    if df_filtered.empty:
        st.warning("Нет данных для выбранных этажей.")
    else:
        # ---- График 1: Углы X и Y ----
        fig1 = go.Figure()

        # Цветовая палитра (качественная, из Plotly)
        colors = px.colors.qualitative.Plotly

        for i, floor in enumerate(sorted(selected_floors)):
            floor_df = df_filtered[df_filtered['Этаж'] == floor].sort_values('Цикл')
            color = colors[i % len(colors)]

            # Углы X – сплошная линия, круглые маркеры
            fig1.add_trace(go.Scatter(
                x=floor_df['Цикл'],
                y=floor_df['αx_abs'],
                mode='lines+markers',
                name=f'Этаж {floor} αx',
                line=dict(color=color, width=2),
                marker=dict(size=8, symbol='circle'),
                legendgroup=f'floor_{floor}',
                legendgrouptitle_text=f'Этаж {floor}'
            ))
            # Углы Y – пунктирная линия, квадратные маркеры
            fig1.add_trace(go.Scatter(
                x=floor_df['Цикл'],
                y=floor_df['αy_abs'],
                mode='lines+markers',
                name=f'Этаж {floor} αy',
                line=dict(color=color, width=2, dash='dot'),
                marker=dict(size=8, symbol='square'),
                legendgroup=f'floor_{floor}',
                showlegend=False  # скрываем дублирующую запись в легенде
            ))

        # Добавляем аннотации максимальных значений (по X и Y)
        max_x = df_filtered.loc[df_filtered['αx_abs'].idxmax()]
        max_y = df_filtered.loc[df_filtered['αy_abs'].idxmax()]
        fig1.add_annotation(
            x=max_x['Цикл'], y=max_x['αx_abs'],
            text=f"max αx = {max_x['αx_abs']:.3f}° (эт.{max_x['Этаж']})",
            showarrow=True, arrowhead=2, ax=0, ay=-40,
            font=dict(color='darkblue', size=11)
        )
        fig1.add_annotation(
            x=max_y['Цикл'], y=max_y['αy_abs'],
            text=f"max αy = {max_y['αy_abs']:.3f}° (эт.{max_y['Этаж']})",
            showarrow=True, arrowhead=2, ax=0, ay=40,
            font=dict(color='darkred', size=11)
        )

        fig1.update_layout(
            title="Абсолютные углы наклона по осям X и Y",
            xaxis_title="Цикл (дата)",
            yaxis_title="Угол, °",
            template="plotly_white",
            hovermode="x unified",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                title="",
                font=dict(size=10)
            ),
            margin=dict(l=40, r=40, t=80, b=40)
        )
        st.plotly_chart(fig1, use_container_width=True)

        # ---- График 2: Сравнение с осадками (если есть) ----
        df_sett_angles = st.session_state.get('df_sett_angles', None)
        if df_sett_angles is not None and not df_sett_angles.empty:
            st.subheader("📊 Сравнение углов по осадкам и наклономеру (этаж 5)")

            # Берём только этаж 5 для сравнения (или другой, если нужно)
            floor_for_compare = 5
            if floor_for_compare not in df_incl['Этаж'].unique():
                floor_for_compare = df_incl['Этаж'].min()
            incl_compare = df_incl[df_incl['Этаж'] == floor_for_compare].copy()

            merged = pd.merge(incl_compare, df_sett_angles, on='Цикл', how='inner')
            if merged.empty:
                st.warning(f"Нет общих циклов для сравнения с осадками (этаж {floor_for_compare}). Показываем последние значения.")
                last_incl = incl_compare.iloc[-1] if not incl_compare.empty else None
                last_sett = df_sett_angles.iloc[-1] if not df_sett_angles.empty else None
                if last_incl is not None and last_sett is not None:
                    st.write("**Последний цикл наклономера (этаж {})**:".format(floor_for_compare))
                    st.dataframe(last_incl[['Цикл', 'αx_abs', 'αy_abs']].to_frame().T)
                    st.write("**Последний цикл осадок**:")
                    st.dataframe(last_sett[['Цикл', 'a_град', 'b_град']].to_frame().T)
            else:
                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(
                    x=merged['Цикл'],
                    y=merged['αx_abs'],
                    mode='lines+markers',
                    name='αx (наклономер)',
                    line=dict(color='#1f77b4', width=2),
                    marker=dict(size=8)
                ))
                fig2.add_trace(go.Scatter(
                    x=merged['Цикл'],
                    y=merged['a_град'],
                    mode='lines+markers',
                    name='a (осадки)',
                    line=dict(color='#ff7f0e', width=2, dash='dash'),
                    marker=dict(size=8, symbol='diamond')
                ))
                fig2.add_trace(go.Scatter(
                    x=merged['Цикл'],
                    y=merged['αy_abs'],
                    mode='lines+markers',
                    name='αy (наклономер)',
                    line=dict(color='#2ca02c', width=2),
                    marker=dict(size=8)
                ))
                fig2.add_trace(go.Scatter(
                    x=merged['Цикл'],
                    y=merged['b_град'],
                    mode='lines+markers',
                    name='b (осадки)',
                    line=dict(color='#d62728', width=2, dash='dash'),
                    marker=dict(size=8, symbol='diamond')
                ))

                fig2.update_layout(
                    title=f"Сравнение углов (этаж {floor_for_compare})",
                    xaxis_title="Цикл",
                    yaxis_title="Угол, °",
                    template="plotly_white",
                    hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig2, use_container_width=True)

        # ---- График 3: Профиль смещений (последний цикл) ----
        last_cycle = df_incl['Цикл'].iloc[-1]
        profile = df_incl[df_incl['Цикл'] == last_cycle].sort_values('Этаж')
        if not profile.empty:
            st.subheader(f"📊 Профиль смещений на цикле {last_cycle}")

            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(
                x=profile['Смещение X'],
                y=profile['Этаж'],
                mode='lines+markers',
                name='Смещение X',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=12, symbol='circle', color='#1f77b4')
            ))
            fig3.add_trace(go.Scatter(
                x=profile['Смещение Y'],
                y=profile['Этаж'],
                mode='lines+markers',
                name='Смещение Y',
                line=dict(color='#ff7f0e', width=3),
                marker=dict(size=12, symbol='square', color='#ff7f0e')
            ))

            # Добавляем вертикальную линию x=0
            fig3.add_vline(x=0, line_width=1, line_dash="dash", line_color="grey", opacity=0.5)

            # Добавляем аннотации значений для каждого этажа
            for _, row in profile.iterrows():
                fig3.add_annotation(
                    x=row['Смещение X'], y=row['Этаж'],
                    text=f"{row['Смещение X']:.3f}",
                    showarrow=False,
                    font=dict(size=9, color='#1f77b4'),
                    xanchor='left', xshift=10
                )
                fig3.add_annotation(
                    x=row['Смещение Y'], y=row['Этаж'],
                    text=f"{row['Смещение Y']:.3f}",
                    showarrow=False,
                    font=dict(size=9, color='#ff7f0e'),
                    xanchor='left', xshift=10
                )

            fig3.update_layout(
                title=f"Профиль смещений (цикл {last_cycle})",
                xaxis_title="Смещение, м",
                yaxis_title="Этаж",
                template="plotly_white",
                yaxis=dict(autorange="reversed", dtick=1),
                xaxis=dict(zeroline=True, zerolinewidth=1, zerolinecolor='grey'),
                hovermode="y unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig3, use_container_width=True)
