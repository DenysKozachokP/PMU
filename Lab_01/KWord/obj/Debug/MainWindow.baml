   M S B A M L   `   `   `  ����e  aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089wTclr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.RibbonSystem.Windows.Controls.Ribbon  5   	   @ <KWord, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:KWordKWord 5   	      KWord.MainWindow   5      _ribbonTclr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.Ribbon  5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   %localclr-namespace:KWord     5   	     � Title$  
MainWindow��5	   	   $	��450��6   $	��800��6)   .��5      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5      $
��Auto �6   .+�5   
   �� -   5         ��Name$ HeaderTextBox��5      $��1��6)   $)��#Верхній колонтитул��66   $��10q�6P   $
��Bold"�6\   $��Center=�6n   �� 5   
   -   5         ��Name$
 doc1��5      $��2��6$   $��10q�61   �� 5   
   -   5        $ FooterTextBox��5      $��3��6)   $'��!Нижній колонтитул��66   $��10q�6O   $
��Bold"�6[   $��Center=�6m   +   %System.Windows.Controls.Ribbon.Ribbon  5   
   $��0��6   $��1,1,1,1q�6      HelpPaneContent 5      1   +System.Windows.Controls.Ribbon.RibbonButton  5       -   5           SmallImageSource$! /images/HelpApplication.png��5    &   6U      ApplicationMenu 5"      :   4System.Windows.Controls.Ribbon.RibbonApplicationMenu  5$      5   /System.Windows.Controls.Ribbon.RibbonMenuButton   SmallImageSource$ \images\iconhome.png��6(   .@�5%      >   8System.Windows.Controls.Ribbon.RibbonApplicationMenuItem  $��Create��60   3   -System.Windows.Controls.Ribbon.RibbonMenuItem   ImageSource$ \images\Document.png��6@   	   KeyTip$	 C��6c     5&      -   5        $
��Open��5&   0   $ \images\OpenFile.png��6>   $	 O��6a   6l     5'      $��Print��60   $ \images\Print.png��6?   $	 P��6_     5(      $��
Properties��60   $ \images\Property.png��6D     5)      -   5        $��Close��5)   0   $ \images\Close.png��6?   $	 C��6_   6j   5*      
   QuickAccessToolBar
 5,      =   7System.Windows.Controls.Ribbon.RibbonQuickAccessToolBar  5-      .@�5.        -   5        $ \images\SaveAll.png��5.   #   6J   6   0System.Windows.Controls.Ribbon.RibbonSplitButton  5/      -   5          Name$
 Undo��5/   (   $ \images\Undo.png��66   .	  (System.Windows.Input.ApplicationCommands 	  Undo   Command# Z 6Z      CommandTarget 6�   ��  �� ElementName$ _richTextBox��.@�50      8
   2System.Windows.Controls.Ribbon.RibbonSplitMenuItem
  $��Undo1��6.   6?   51        52      -	   5        $
 Redo��52   (   $ \images\Redo.png��66    	  Redo# Z 6Z    6�   �� $ _richTextBox��.@�53      
  $��Redo1��6.   6?   54      55      .@�57      .   (System.Windows.Controls.Ribbon.RibbonTab  $
��Home��6   .@�59      0   *System.Windows.Controls.Ribbon.RibbonGroup  $��	Clipboard��6   .@�5:           LargeImageSource$ \Images\paste.png��6'      Label$ Paste��6L      KeyTip$ V��6Z   .@�5;        $ \Images\Paste.png��6)   $��Keep Text Only��6I   $	 T��6a    	  Paste m� Command# Z 6m    m� CommandTarget 6�   �� $ _richTextBox��  5<      $ \Images\Paste.png��6)   $��Paste Special...��6I   $	 S��6c   5=        5>      $ \Images\Cut.png��6#      Label$	 Cut��6F      KeyTip$ X��6R   
 	  Cut#��Z 6^   ��6�   �� $ _richTextBox��  5?      $ \Images\Copy.png��6#   $
 Copy��6G   $ C��6T    	  Copy#��Z 6`   ��6�   �� $ _richTextBox��5@        5B      -
   5          Name$ fonts��5B      $��Fonts��6-   $
��Auto��6<   .@�5C      7   1System.Windows.Controls.Ribbon.RibbonControlGroup  .@�5D      �� -   5         �� Text$ Select Font...��5D   �    �� 
IsEditable . ��6#   ��    System.Windows.Media.Fonts   SystemFontFamilies �� Source# Z 6�   �� 5E      -   5          ��Name$  	_fontSize��5E   O   $ Size...��6b    . 6�   5F        5G      .@�5H        $ Images\fonts\Bold.png��6'   $ B��6P   .  (System.Windows.Documents.EditingCommands!   
ToggleBold#��Z! 6[   $��20��6�   $��20��6�   ��6�   �� $ _richTextBox��6�     5I      $ Images\fonts\Italic.png��6'   $ I��6R   "   ToggleItalic#��Z" 6^   $��20��6�   $��38��6�   ��6�   �� $ _richTextBox��6�     5J      $  Images\fonts\Underline.png��6'   $ U��6U   #   ToggleUnderline#��Z# 6b   $��15��6�   $��26��6�   ��6�   �� $ _richTextBox��6�   5K        5L      $! \Images\fonts\FontColor.png��6'   $ 
Font Color��6V   $ C��6i   .@�5M        $��Чорний��6)   
���   �$ m� CommandParameter$$ Black��6s    6L   �� *ChangeColorCommand+6s     5N      $��
Синій��6)   
����  �$
$ Blue��6q    6J   �� *ChangeColorCommand+6q     5O      $��Жовтий��6)   
��� ���$$ Yellow��6t    6M   �� *ChangeColorCommand+6t     5P      $��Червоний��6)   
���  ��$	$ Red��6s    6L   �� *ChangeColorCommand+6s     5Q      $��Зелений��6)   
��� � �$$ Green��6t    6M   �� *ChangeColorCommand+6t     5R      $��
Білий��6)   
�������$$ White��6r    6K   �� *ChangeColorCommand+6r   5S      5U        5V      -   5        $ 	paragraph��5V      $��	Paragraph��61   .@�5W        .@�5X        $! \Images\align\AlignLeft.png��6(   %   	AlignLeft#��Z% 6W   ��6�   �� $ _richTextBox��  5Y      $" \Images\align\AlignRight.png��6'   &   
AlignRight#��Z& 6W   ��6�   �� $ _richTextBox��  5Z      $! \Images\align\AlignBoth.png��6(   '   AlignJustify#��Z' 6W   ��6�   �� $ _richTextBox��5[        5\      .@�5]        $" \Images\align\BulletList.png��6'   5^      5_      5`        5a      $��Insert��6   $��-8,0,8,0q�6(   ( � UseLayoutRounding( . $
��Auto=�6S   .@�5b        -   5        $ objects��5b      $��Objects��6/   .@�5c        )   LargeImageSource$) \Images\insert\Table.png��6#   $ Table��6O     5d      $$ \Images\insert\Application.png��6#     5e      $" \Images\insert\Rectangle.png��6#     5f      $" \Images\insert\LineChart.png��6#     5g      $ \Images\insert\Image.png��6#   5h        5i      $��Headers and Footers��6   .@�5j        -   5        $) \Images\insert\Image.png��5j   #   $ Edit Header��6O   6c     5k      -   5        $) \Images\insert\Image.png��5k   #   $ Edit Footer��6O   6c   5l      5m      5n      5o      5s      