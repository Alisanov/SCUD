   M S B A M L   `   `   `  ����?  ;SCUD, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SCUDSCUD  5   	       SCUD.MyMessageBox   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   $localclr-namespace:SCUD     5   	   -   5          �Name$  
MessageBox��5
   v    � Title$ Message Box��5   	   $	��200��6   $	��400��6*    � WindowStartupLocation$ CenterScreen=�66      Main#��[  6[    � WindowState$ Normal=�6y    � ShowInTaskbar .   � AllowsTransparency .  � 
ResizeMode$ NoResize=�5
   	    � WindowStyle$
 None=�6    J� Opacity$
 0.01Q�6g   	 � Triggers	 5      :� 5   
   
 :� RoutedEvent$
 Loaded��6   .[�5      �� .��5      �� .��5      W� 	 W� To$
 0.95X�5   !    g� Duration$ 0:0:0.4D�5   !    �� TargetProperty$ Opacity�5   !   5      5      5      .��5#      � ���#9A9A9A.+�5$   
   � $��2q�6   ���#F3FAFF � RowDefinitions 5%      �� 5&      $��50 �6    �� 5'      $��7* �6    �� 5(      $
��Auto �6    .+�5*      �� -   5        $��Caption��5*      $��10q�6'   $��0��63   �� 5+      -   5        $��MessageContainer��5+      $��WrapWithOverflow=�60   $��10q�6P   $��1��6\   �� 5,      -   5        $��ButtonContainer��5,      $��2��60    �� Orientation$ 
Horizontal=�6=   $��Right=�5-      $��
0,10,10,10q�61    � 	Resources 5.       �� 
TargetType �� Property �� Value��%�   '��       �� 5/      " ��6   .�50      ��  � ��$ 25��66   �� 51       � ��$ 75��65   �� 52       � ��$ 10,0,0,0q�63   53      54      55      56      57      58      