   M S B A M L   `   `   `  ���� ?  ;SCUD, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SCUDSCUD      SCUD.MyMessageBox   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     $localclr-namespace:SCUD     -     �Name$  
MessageBox�� � Title$ Message Box��$	��200��$	��400�� � WindowStartupLocation$ CenterScreen=�   Main#��[   � WindowState$ Normal=� � ShowInTaskbar .   � AllowsTransparency .  � 
ResizeMode$ NoResize=� � WindowStyle$
 None=� J� Opacity$
 0.01Q�	 � Triggers	 :� 
 :� RoutedEvent$
 Loaded��.[��� .���� .��W� 	 W� To$
 0.95X� g� Duration$ 0:0:0.4D� �� TargetProperty$ Opacity�.��� ���#9A9A9A.+�� $��2q����#F3FAFF � RowDefinitions �� $��50 ��� $��7* ��� $
��Auto �.+��� -   $��Caption��$��10q�$��0���� -   $��MessageContainer��$��WrapWithOverflow=�$��10q�$��1���� -   $��ButtonContainer��$��2�� �� Orientation$ 
Horizontal=�$��Right=�$��
0,10,10,10q� � 	Resources  �� 
TargetType �� Property �� Value��%_   '��       �� " ��.���  � ��$ 25����  � ��$ 75����  � ��$ 10,0,0,0q�