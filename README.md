<?php
	use PhpOffice\PhpWord\Style\Font;
	
	include_once 'Sample_Header.php';
	
	
	echo date('H:i:s') , ' Create new PhpWord object' , EOL;
	
	$languageEnGb = new \PhpOffice\PhpWord\Style\Language(\PhpOffice\PhpWord\Style\Language::EN_GB);
	
	$phpWord = new \PhpOffice\PhpWord\PhpWord();
	$phpWord->getSettings()->setThemeFontLang($languageEnGb);
	
	$fontStyleName = 'rStyle';
	$phpWord->addFontStyle($fontStyleName, array('bold' => true, 'italic' => true, 'size' => 16, 'allCaps' => true, 'doubleStrikethrough' => true));
	
	$paragraphStyleName = 'pStyle';
	$phpWord->addParagraphStyle($paragraphStyleName, array('alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 100));
	
	$phpWord->addTitleStyle(1, array('bold' => true), array('spaceAfter' => 240));
	
	
	$section = $phpWord->addSection();
	
	
	$section->addTitle('Welcome to PhpWord', 1);
	$section->addText('Hello World!');
	
	// $pStyle = new Font();
	// $pStyle->setLang()
	$section->addText('Ce texte-ci est en français.', array('lang' => \PhpOffice\PhpWord\Style\Language::FR_BE));
	
	
	$section->addTextBreak(2);
	
	
	$section->addText('I am styled by a font style definition.', $fontStyleName);
	$section->addText('I am styled by a paragraph style definition.', null, $paragraphStyleName);
	$section->addText('I am styled by both font and paragraph style.', $fontStyleName, $paragraphStyleName);
	
	$section->addTextBreak();
	
	
	$fontStyle['name'] = 'Times New Roman';
	$fontStyle['size'] = 20;
	
	$textrun = $section->addTextRun();
	$textrun->addText('I am inline styled ', $fontStyle);
	$textrun->addText('with ');
	$textrun->addText('color', array('color' => '996699'));
	$textrun->addText(', ');
	$textrun->addText('bold', array('bold' => true));
	$textrun->addText(', ');
	$textrun->addText('italic', array('italic' => true));
	$textrun->addText(', ');
	$textrun->addText('underline', array('underline' => 'dash'));
	$textrun->addText(', ');
	$textrun->addText('strikethrough', array('strikethrough' => true));
	$textrun->addText(', ');
	$textrun->addText('doubleStrikethrough', array('doubleStrikethrough' => true));
	$textrun->addText(', ');
	$textrun->addText('superScript', array('superScript' => true));
	$textrun->addText(', ');
	$textrun->addText('subScript', array('subScript' => true));
	$textrun->addText(', ');
	$textrun->addText('smallCaps', array('smallCaps' => true));
	$textrun->addText(', ');
	$textrun->addText('allCaps', array('allCaps' => true));
	$textrun->addText(', ');
	$textrun->addText('fgColor', array('fgColor' => 'yellow'));
	$textrun->addText(', ');
	$textrun->addText('scale', array('scale' => 200));
	$textrun->addText(', ');
	$textrun->addText('spacing', array('spacing' => 120));
	$textrun->addText(', ');
	$textrun->addText('kerning', array('kerning' => 10));
	$textrun->addText('. ');
	
	
	$section->addLink('PHPWord');
	$section->addTextBreak();
	
	// Image
	$section->addImage('resources/_earth.jpg', array('width'=>18, 'height'=>18));
	
	// Save file
	echo write($phpWord, basename(__FILE__, '.php'), $writers);
	if (!CLI) {
	    include_once 'Sample_Footer.php';
	}
<?php
	include_once 'Sample_Header.php';
	
	
	echo date('H:i:s'), ' Create new PhpWord object', EOL;
	$phpWord = new \PhpOffice\PhpWord\PhpWord();
	
	
	$paragraphStyleName = 'pStyle';
	$phpWord->addParagraphStyle($paragraphStyleName, array('spacing' => 100));
	
	$boldFontStyleName = 'BoldText';
	$phpWord->addFontStyle($boldFontStyleName, array('bold' => true));
	
	$coloredFontStyleName = 'ColoredText';
	$phpWord->addFontStyle($coloredFontStyleName, array('color' => 'FF8080', 'bgColor' => 'FFFFCC'));
	
	$linkFontStyleName = 'NLink';
	$phpWord->addLinkStyle($linkFontStyleName, array('color' => '0000FF', 'underline' => \PhpOffice\PhpWord\Style\Font::UNDERLINE_SINGLE));
	
	
	$section = $phpWord->addSection();
	
	
	$textrun = $section->addTextRun($paragraphStyleName);
	$textrun->addText('Each textrun can contain native text, link elements or an image.');
	$textrun->addText(' No break is placed after adding an element.', $boldFontStyleName);
	$textrun->addText(' Both ');
	$textrun->addText('superscript', array('superScript' => true));
	$textrun->addText(' and ');
	$textrun->addText('subscript', array('subScript' => true));
	$textrun->addText(' are also available.');
	$textrun->addText(' All elements are placed inside a paragraph with the optionally given paragraph style.', $coloredFontStyleName);
	$textrun->addText(' Sample Link: ');
	$textrun->addLink('PHPWord', $linkFontStyleName);
	$textrun->addText(' Sample Image: ');
	$textrun->addImage('resources/_earth.jpg', array('width' => 18, 'height' => 18));
	$textrun->addText(' Sample Object: ');
	$textrun->addObject('resources/_sheet.xls');
	$textrun->addText(' Here is some more text. ');
	
	$textrun = $section->addTextRun();
	$textrun->addText('This text is not visible.', array('hidden' => true));
	
	
	echo write($phpWord, basename(__FILE__, '.php'), $writers);
	if (!CLI) {
	    include_once 'Sample_Footer.php';
	}


