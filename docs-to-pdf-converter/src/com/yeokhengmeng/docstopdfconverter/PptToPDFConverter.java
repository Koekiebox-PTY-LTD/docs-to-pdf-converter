package com.yeokhengmeng.docstopdfconverter;

//import org.apache.poi.sl.usermodel.Slide;
//import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.SlideShow;

import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShowFactory;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

//import org.apache.poi.hslf.model.Slide;
//import org.apache.poi.hslf.usermodel.SlideShow;

public class PptToPDFConverter extends PptxToPDFConverter {

	private java.util.List<Slide> slides;
	
	public PptToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}


	@Override	
	protected Dimension processSlides() throws IOException{

		SlideShow ppt = SlideShowFactory.create(this.inStream);
		Dimension dimension = ppt.getPageSize();
		slides = ppt.getSlides();
		return dimension;
	}
	
	@Override
	protected int getNumSlides(){
		return slides.size();
	}
	
	@Override
	protected void drawOntoThisGraphic(int index, Graphics2D graphics){
		slides.get(index).draw(graphics);
	}
	
	@Override
	protected Color getSlideBGColor(int index){

		PaintStyle paintStyle = null;

		if(this.slides.get(index).getBackground() != null &&
				this.slides.get(index).getBackground().getFillStyle() != null)
		{
			paintStyle = this.slides.get(index).getBackground().getFillStyle().getPaint();
		}

		if(paintStyle != null && paintStyle instanceof PaintStyle.SolidPaint)
		{
			return ((PaintStyle.SolidPaint)paintStyle).getSolidColor().getColor();
		}

		return null;
		//return this.slides.get(index).getBackground().getFillStyle().getPaint().get;
	}

}
