   M S B A M L   `   `   `  ���� B  >xls2sql, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:xls2sqlxls2sql      xls2sql.MainWindow   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     'localclr-namespace:xls2sql     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006       � Title$&   xls2sql 2.2 (by Daniel Haliscak)��$	��800��$	��450�� � 
ResizeMode$ CanResizeWithGrip=�.��� .+�]� $��10,10,9.6,10q�.+�� $��0,0,0,10q�$
��left=�$	��Top=�.+��� -    ��Name$ Database��$��Center=�$��Database���� -   $��txtDatabaseName��$	��200��$��27��$��5,0,5,0q�$
��Left=�$��Center=�$
��Left=�$��Center=�$
{�True�� �� 	MaxLength$	 100�� �� TextWrapping$ NoWrap=��� -   $ Table��$��Center=�$��Table���� -   $��txtTableName��$	��200��$��27��$��5,0,5,0q�$
��Left=�$��Center=�$
��Left=�$��Center=�$
{�True��$	 100��$ NoWrap=��� -   $��ckbCreateTable��$��	10,0,10,0q�$��Center=�$��LeftToRight=� d� 	IsChecked$
 TrueY�.���� $��LeftToRight=�$"��Generate CREATE TABLE Script��� $��0,0,0,10q�$
��left=�$	��Top=�.+��� -   $ FirstColumn��$��Center=�$��First Column���� -   $��cmbFirstColumn��$	��150��$��5q�$��5q� �� SelectedIndex$ 0��.@��� $
��None���� $��Id (int, IDENTITY 1,1)���� $��Id (NewSequentialId())���� -   $ 	Separator��$��Center=�$��	Separator���� -	   $��txtSeparator��$	��100��$��27��$��5,0,5,0q�$
��Left=�$��Center=�$
��Left=�$��Center=�$
{�True�� J� RenderTransformOrigin$	 0,0&�$
��1000��$ NoWrap=� �� 	UndoLimit$	 102���� -
   $��ckbTrimWhiteSpaces��$��	10,0,10,0q�$��Center=�$��LeftToRight=�$
 TrueY�.���� $��LeftToRight=�$��Trim White Spaces���� -   $��ckbPrefferNulls��$��	10,0,10,0q�$��Center=�$��LeftToRight=�$
 TrueY�.���� $��LeftToRight=�$��Preffer Nulls��� $	��Top=�.+��� -   $
 Path��$��Center=�$
��Path���� -   $��txtFilepath��$	��371��$��27��$
��Left=�$��Center=�$
��Left=�$��Center=�	 J� 	AllowDrop$
	 True��
 �� 
IsReadOnly$

 True��$ NoWrap=�.������  �� Color# � ��.���� -   $��btnOpenFile��$��0,5,5,5q�$��5q�.��	Open File�� -   $��cmbWorkbook��$	��100��$��5q�$��5q�$y�	Collapsed=��� -   $��
btnExecute��$��5q�$��5q�.��
Generate�  � RowDefinitions �� $
��Auto ��� �� $
��Auto �.+��� $��0��$��Center=�$��Output���� -   $��	txtEditor��$��1��$��0,0,0,0q�$ �� HorizontalScrollBarVisibility$ Disabled=�$
��True��$
 Wrap=�" �� VerticalScrollBarVisibility$ Visible=��� -   $��	txtStatus��$��2��$
��Left=�$�� ��