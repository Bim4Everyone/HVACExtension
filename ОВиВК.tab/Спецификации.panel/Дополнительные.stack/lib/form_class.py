#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from pyrevit import forms


class SelectParametersForm:
    layout = '''
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Выбор параметров"
        Width="300" Height="190"
        ShowInTaskbar="False" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen">

        <Grid Margin="5">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Label Grid.Row="0" Content="Выберите откуда копируем:"/>
            <ComboBox Grid.Row="1" Name="combobox1" Width="250"/>

            <Label Grid.Row="2" Content="Выберите куда копируем:"/>
            <ComboBox Grid.Row="3" Name="combobox2" Width="250"/>

            <Button Grid.Row="4" Name="okButton" Content="OK" Width="75" Height="25" HorizontalAlignment="Right" Margin="0,10,0,0"/>
        </Grid>
    </Window>
    '''

    def __init__(self, table_column_names, table_params):
        self.table_column_names = table_column_names
        self.table_params = table_params
        self.result = (None, None)

    def show_form(self):
        w = forms.WPFWindow(self.layout, literal_string=True)

        # Переменная для защиты от бесконечных обновлений
        self.is_updating = False

        # Функция для заполнения комбобоксов с разными начальными значениями
        def initialize_comboboxes():
            if len(self.table_column_names) < 2:
                w.combobox1.Items.Add(self.table_column_names[0])
                w.combobox2.Items.Add(self.table_column_names[0])
                w.combobox1.SelectedIndex = 0
                w.combobox2.SelectedIndex = 0
                return

            w.combobox1.Items.Clear()
            w.combobox2.Items.Clear()

            for value in self.table_column_names:
                w.combobox1.Items.Add(value)

            for value in self.table_params:
                w.combobox2.Items.Add(value)

            w.combobox1.SelectedIndex = 0
            w.combobox2.SelectedIndex = 1
            update_combobox2_options()

        # Обновление списка доступных значений во втором комбобоксе
        def update_combobox2_options():
            if self.is_updating:
                return

            self.is_updating = True

            selected1 = w.combobox1.SelectedItem
            selected2 = w.combobox2.SelectedItem

            w.combobox2.Items.Clear()
            for value in self.table_params:
                if value != selected1:
                    w.combobox2.Items.Add(value)

            # Если предыдущий выбранный элемент отсутствует, выбираем первый
            if selected2 in w.combobox2.Items:
                w.combobox2.SelectedItem = selected2
            else:
                w.combobox2.SelectedIndex = 0

            self.is_updating = False

        # Обновление списка доступных значений в первом комбобоксе
        def update_combobox1_options():
            if self.is_updating:
                return

            self.is_updating = True

            selected1 = w.combobox1.SelectedItem
            selected2 = w.combobox2.SelectedItem

            w.combobox1.Items.Clear()
            for value in self.table_column_names:
                if value != selected2:
                    w.combobox1.Items.Add(value)

            if selected1 in w.combobox1.Items:
                w.combobox1.SelectedItem = selected1
            else:
                w.combobox1.SelectedIndex = 0

            self.is_updating = False

        # Обработчики событий изменения выбора
        def on_combobox1_changed(sender, args):
            update_combobox2_options()

        def on_combobox2_changed(sender, args):
            update_combobox1_options()

        w.combobox1.SelectionChanged += on_combobox1_changed
        w.combobox2.SelectionChanged += on_combobox2_changed

        # Функция обработки нажатия кнопки
        def on_ok(sender, args):
            self.result = (w.combobox1.SelectedItem, w.combobox2.SelectedItem)
            w.Close()

        # Назначаем обработчик кнопке
        w.okButton.Click += on_ok

        # Инициализация и запуск окна
        initialize_comboboxes()
        w.show_dialog()

        return self.result