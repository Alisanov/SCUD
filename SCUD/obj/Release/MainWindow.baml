   M S B A M L   `   `   `  ���� ?  ;SCUD, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SCUDSCUD      SCUD.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     $localclr-namespace:SCUD     X TXceed.Wpf.Toolkit, Version=4.0.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f44xctk)http://schemas.xceed.com/wpf/xaml/toolkit  -     �Name$  window�� � Title$N HСистема контроля и управления доступом��   Main#��[   � WindowStartupLocation$ CenterScreen=� � WindowState$ Normal=�$	��400��$	��400��.��� .+�� -    �Name$ loginWindow��$y�	Collapsed=�.+��� .+��� $��18(�$��Center=�$��20 100 20 0q�$Z��TВход в систему контроля и управления доступом�� �� TextAlignment$ Center=�$
��Wrap=�$
��Bold"��� $��
Логин���� -    ��Name$ login��$��admin��  ShadowTextBox#��[ �� $��Пароль��G� -    G�Name$ password�� G� Password$ admin��  ShadowPasswordBox#��[ �� -   	 ��Name$	 infoBar��$��0 5 0 0q�$��0��$��14(�
���  ��$
��Bold"�$��Center=��� -   $��	20 5 20 0q�$��
Войти��  ShadowButton#��[ � -   $ panel��$y�Visible=�
 � ColumnDefinitions
 �� �� ��  � RowDefinitions �� $��70 ��� �� $��50 �.+�� $��3��$��1q�
 �� �� �� .+��� -    ��Name$ Canvas��
������ $��25��$	��Top=�.+��� -	   $	 mainText��$��Bottom=�$��16(�$
��Bold"� � Triggers :�  :� RoutedEvent$ TextBlock.Loaded��[��� .���� .��W�  �� TargetProperty$ (Canvas.Right)� W� From$ 1000 X�	 W� To$ -1500X� g� Duration$ 0:0:15D� g� RepeatBehavior$ Forever���� -
   $��1q�$<��6Заполнить список сотрудников��#��[ �� -   $��1q�$��1��$/��)Заполнить справочники��#��[ �� $��1q�$��2��$��Настройки��#��[ � $��5q�$��1�� �� $��30 ��� $��30 ��� $	��70* �.+��� $��Center=�$
��Wrap=�$)��#Список сотрудников��$��14(�$
��Bold"�� $��1��
 �� $	��30* ��� $	��10* �.+��� -   $ findText��#��[ $��5 0 5 0q��� -   $��
Поиск��#��[ $��1���� $��2��$
��Auto=�.���� -    ��Name$ employeeListPanel��  listEmployeePanel#��[ ���#F3FAFF$��0q�$��Disabled=�$
��Auto=� %� 
ScrollUnit$ Pixel=����� .R�� �� $��5q�$��1��$��1��.+��� $��Сотрудник��.���� .+��� -    ��Name$
 name��$��Black"�$��15(�$�� ���� -    ��Name$ 	timeInput��.+�� $��2q�
 �� �� .+��� $
��Left=�$��	Дата:��(  "System.Windows.Controls.DatePicker  -     Name$ dateEmployee��$��1��$��14(�$
��Bold"�� $��2q�
 �� �� .+��� $
��Left=�$ ��Время прихода:��"  Xceed.Wpf.Toolkit.TimePicker  -     Name$ 	timeStart��$��14(�$��1��$
��Bold"�,  &Xceed.Wpf.Toolkit.Primitives.InputBase   Text$
 8:00��� $��2q�
 �� �� .+��� $
��Left=�$��Время ухода:��  -   $ timeEnd��$��14(�$��1��$
��Bold"�$ 18:00���� -    ��Name$ 
checkInput��$5��/Заполнять автоматический���� -   #��[ $��3q�$%��Сотрудник пришел���� -   #��[ $��3q�$!��Сотрудник ушел��� $��1��$��2�� �� $��50 ��� $	��30* �.+��� -   $	 employeeActualText��$��2q�$
��Wrap=�$��14(�$
��Bold"��� $��1��.���� -   $ employeeListActual��  listEmployeePanelTime#��[ ���#F3FAFF$��0q�$
��Auto=�$ Pixel=�� $��2��$��3��
 �� �� �� .+��� -   #��[ $��3q�$)��#Сформировать отчет���� -   $��1��#��[ $��3q�$N��HСформировать отчет по всем сотрудникам���� -   $��2��#��[ $��3q�$ ��Выбрать файл...��� -   $ employeeArea��$��10q�$y�	Collapsed=�
 �� $��30 ��� ��  �� �� $��50 �.+��� -   $	��←��$	��200��#��[ $��1q�$��3��$	��Top=��� $��1��$
��Auto=�.���� -   $ employeeList��  listEmployee#��[ ���#F3FAFF$��0q�! �� IsDeferredScrollingEnabled . $ Pixel=��� -   $��Удалить��$��1��$��1��#��[ �� $+��%Добавить сотрудника��$��2��$��2��.���� .+��� $��	20 10 0 0q�$��Имя��$
��Left=��� -    $ nameEmployee��#��[ �� $
��20 0q�$��Фамилия��$
��Left=��� -!   $ lastnameEmployee��#��[ �� $
��20 0q�$��Должность��$
��Left=��� -"    ��Name$ positionComboBox��$
��20 0q��� $
��20 0q�$#��Рабочая область��$
��Left=��� -#   $ workareaComboBox��$
��20 0q��� $
��20 0q�$��
Отдел��$
��Left=��� -$   $ departamentComboBox��$
��20 0q��� -%   $	 infoBarEmployee��$��0 5 0 0q�$��0��$��14(�
���  ��$
��Bold"�$��Center=��� -&   #��[ $��20 10q�$��Добавить��� -'   $ directoresArea��$��5q�$y�	Collapsed=�
 �� $��30 ��� �� ��  �� �� $��50 �.+��� -(   $	��←��$	��200��#��[ $��1q�$��2��$	��Top=��� $��1��$
��Auto=�$��2��.���� -)   $��1q�$��0��.+��� $��Center=�$
��Wrap=�$!��Список отделов��$��14(�$
��Bold"��� -*   $ departamentList��  listDirectores#��[  . $ Pixel=��� -+   $ departamentText��$��15(�$
��Bold"�$��25���� -,   $��1q�#��[ $��Добавить���� $��2��$
��Auto=�$��2��.���� --   $��1q�$��1��$��0��.+��� $��Center=�$
��Wrap=�$2��,Список рабочих областей��$��14(�$
��Bold"��� -.   $ workareaList��#��[  . $ Pixel=��� -/   $ workareaText��$��15(�$
��Bold"�$��25���� -0   #��[ $��1q�$��Добавить���� $��3��$
��Auto=�$��2��.���� -1   $��2��$��0��$��1q�.+��� $��Center=�$
��Wrap=�$'��!Список должностей��$��14(�$
��Bold"��� -2   $ positionList��#��[  . $ Pixel=��� -3   $ positionText��$��15(�$
��Bold"�$��25���� -4   #��[ $��1q�$��Добавить���� -5   $��2��$��4��$��Удалить��#��[ $
��0,10q�