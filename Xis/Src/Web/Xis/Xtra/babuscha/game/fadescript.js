	function makearray(n)
	{
		this.length = n;
		for(var i = 1; i <= n; i++)
		{
			this[i] = 0;
		}
		return this;
	}
	hexa = new makearray(16);
	for(var i = 0; i < 10; i++) hexa[i] = i;
	hexa[10]="a"; hexa[11]="b"; hexa[12]="c";
	hexa[13]="d"; hexa[14]="e"; hexa[15]="f";
	
	function hex(i)
	{
		if (i < 0) return "00";
		else if (i >255) return "ff";
		else return "" + hexa[Math.floor(i/16)] + hexa[i%16];
	}
	
	function setbgColor(r, g, b)
	{
		var hr = hex(r); var hg = hex(g); var hb = hex(b);
		document.bgColor = "#"+hr+hg+hb;
	}
	
	function fade(sr, sg, sb, er, eg, eb, step)
	{
		for(var i = 0; i <= step; i++)
		{
			setbgColor(Math.floor(sr * ((step-i)/step) + er * (i/step)),
			Math.floor(sg * ((step-i)/step) + eg * (i/step)),
			Math.floor(sb * ((step-i)/step) + eb * (i/step)));
		}
	}
	/* Usage:
	*   fade(inr, ing, inb, outr, outg, outb, step);
	* example.
	*   fade(0,0,0, 255,255,255, 255);
	* fade from black to white with very slow speed.
	*   fade(255,0,0, 0,0,255, 50);
	*   fade(0xff,0x00,0x00, 0x00,0x00,0xff, 50); // same as above
	* step 2 is very fast and step 255 is very slow.
	*/
	function fadein()
	{
		//fade(255, 255, 255, 0, 0, 0, 100);
		//fade(51, 102, 102, 0, 0, 0, 100);
		//fade(51, 102, 102, 153, 204, 204, 100);
		fade(153, 204, 204, 51, 102, 102, 100);
		
	}

	function fadeout()
	{
		//fade(0, 0, 0, 255, 255, 255, 100);
		//fade(0, 0, 0, 51, 102, 102, 100);
		fade(51, 102, 102, 153, 204, 204, 100);
	}
