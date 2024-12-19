import pandas as pd
import os

current_directory = os.getcwd()

def create_smeta_excel(excel_path = os.path.join(current_directory, 'smeta.xlsx')):

    # Данные о работах и их стоимости
    work_data = {
        'Номер бригады': [1, 1, 1, 2, 2, 3, 4],
        'Тип работы': ['Штукатурка', 'Шпатлёвка', 'Подготовка стены под покраску',
                       'Покраска стен', 'Поклейка обоев', 'Установка натяжных потолков', 'Укладка ламината'],
        'Стоимость за м² (руб.)': [1000, 1500, 2500, 1300, 1700, 2200, 1800]
    }

    # Размеры комнаты
    room_length = 4  # метры
    room_width = 3  # метры
    room_height = 2.7  # метры

    # Площадь поверхностей
    wall_area = 2 * room_height * (room_length + room_width)  # Площадь стен
    ceiling_area = room_length * room_width  # Площадь потолка
    floor_area = room_length * room_width  # Площадь пола

    # Соответствие работы и площади
    work_areas = {
        'Штукатурка': wall_area,
        'Шпатлёвка': wall_area,
        'Подготовка стены под покраску': wall_area,
        'Покраска стен': wall_area,
        'Поклейка обоев': wall_area,
        'Установка натяжных потолков': ceiling_area,
        'Укладка ламината': floor_area
    }

    # Создание таблицы сметы
    smeta_df = pd.DataFrame(work_data)
    smeta_df['Обрабатываемая площадь (м²)'] = smeta_df['Тип работы'].map(work_areas)
    smeta_df['Итоговая стоимость (руб.)'] = smeta_df['Стоимость за м² (руб.)'] * smeta_df['Обрабатываемая площадь (м²)']

    # Итоговая сумма всех работ
    total_cost = smeta_df['Итоговая стоимость (руб.)'].sum()

    # Распределение средств
    foreman_share = total_cost * 0.3  # 30% для прораба
    misc_expenses = total_cost * 0.2  # 20% на прочие расходы
    team_share = total_cost * 0.5  # 50% для бригад

    # Создание таблицы распределения доходов
    distribution_df = pd.DataFrame({
        'Тип распределения средств': ['Общая стоимость работ', 'Прораб (30%)', 'Бригада (50%)', 'Прочие расходы (20%)'],
        'Сумма распределения (руб.)': [total_cost, foreman_share, team_share, misc_expenses]
    })

    # Расчет распределения для каждой бригады
    team_distribution = smeta_df.groupby('Номер бригады')['Итоговая стоимость (руб.)'].sum().reset_index()
    team_distribution['Доля от общей стоимости (%)'] = (team_distribution['Итоговая стоимость (руб.)'] /
                                                        team_distribution['Итоговая стоимость (руб.)'].sum()) * 100
    team_distribution['Чистый доход бригады (руб.)'] = team_share * (
                team_distribution['Итоговая стоимость (руб.)'] / total_cost)

    # Переименовываем столбцы для ясности
    team_distribution.rename(columns={
        'Итоговая стоимость (руб.)': 'Стоимость работ бригады (руб.)'
    }, inplace=True)

    # Запись данных в Excel
    excel_path = os.path.join(current_directory, 'smeta.xlsx')
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        smeta_df.to_excel(writer, index=False, sheet_name='Смета на ремонт')
        distribution_df.to_excel(writer, index=False, sheet_name='Распределение доходов')
        team_distribution.to_excel(writer, index=False, sheet_name='Доход бригад')

    print(f"Смета успешно сохранена в файл: {excel_path}")


if __name__ == '__main__':
    # Указываем путь к файлу, который будет создан
    create_smeta_excel(os.path.join(current_directory, 'smeta.xlsx'))
