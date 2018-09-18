this.Header=function Header()
{
    this.Image('2013_Header.jpg',10,8,190);
    this.SetFont('Helvetica','B',12);
    this.Cell(80);
    this.Ln(38);
    this.Ln(10);
}
this.ChapterTitle=function ChapterTitle(title)
{
    this.SetFont('Helvetica','',12);
    this.SetFillColor(200,220,255);
	this.SetTextColor(000);
    this.Cell(0,6, title ,0,1,'L',true);
	this.Ln(2);
    this.y0=this.GetY();
}
this.ChapterTitle2=function ChapterTitle2(title)
{
    this.SetFont('Helvetica','B',12);
	this.SetFillColor(200);
	this.SetTextColor(000);
    this.Cell(0,6, title ,0,1,'L',true);
	this.Ln(4);
    this.y0=this.GetY();
}
this.GreyTitle=function GreyTitle(title)
{
    this.SetFont('Arial','',7);
    this.SetFillColor(200);
	this.SetTextColor(000);
	this.SetFont('','B'); 
    this.Cell(0,3, title ,0,1,'L',true);
	this.Ln(2);
    this.y0=this.GetY();
}
this.GreyTitle_NS=function GreyTitle_NS(title)
{
    this.SetFont('Arial','',7);
    this.SetFillColor(200);
	this.SetTextColor(000);
	this.SetFont('','B'); 
    this.Cell(0,3, title ,0,1,'L',true);
    this.y0=this.GetY();
}
this.ML=function ML(title)
{
    this.SetFont('Arial','',7);
    this.SetFillColor(200);
	this.SetTextColor(000);
    this.MultiCell(0,3, title ,0,1,'L',true);
    this.y0=this.GetY();
}
this.OrangeTitle=function OrangeTitle(title)
{
    this.SetFont('Helvetica','',10);
    this.SetFillColor(230,200,0);
	this.SetTextColor(000);
    this.Cell(0,4, title ,0,1,'L',true);
    this.y0=this.GetY();
}
this.ChapterBody=function ChapterBody(content)
{
	this.SetFont('Arial','',11);
	this.SetTextColor(0,0,0);
	this.MultiCell(0,5,content);
}
this.Red=function Red(content)
{
	this.SetTextColor(50,210,50);
	this.Cell(0,5,content,0,1);
}
this.Compare=function Compare(status,color)
{
	this.SetTextColor(color);
	this.Cell(0,5,status,0,1);
}
this.Footer=function Footer()
{
    this.SetY(-15);
    this.SetFont('Helvetica','I',8);
    this.Cell(0,10,'Page '+ this.PageNo()+ '/{nb}',0,0,'C');
}
this.TableHeader=function TableHeader(data)
{  
    	this.SetFillColor(230,200,0);
    	this.SetTextColor(0);
    	this.SetLineWidth(.3);
    	this.SetFont('','B'); 
    	this.Cell(35,4,data,0,0,'C',true);
}
this.TablePre=function TablePre()
{
		this.SetFillColor(200);
    	this.SetTextColor(0);
    	this.SetDrawColor(230,200,0);
    	this.SetFont('');
		fill=false;
}
this.TableColumn=function TableColumn(data)
{	
	    this.SetLineWidth(.3);
		this.Cell(35,20,'','LR',0,'C',fill);
		this.MultiCell(35,4,data,'LR',0,'C')
}
this.TablePost=function TablePost()
{
	this.Ln();
    fill=!fill;
}