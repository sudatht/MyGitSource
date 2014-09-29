
	/// Initialization
	var rootFrameIndex = 0;
	var menuFrameIndex = 1;
	var funcFrameIndex = 2;

	//Set correct frame indexes
	SetCorrectFrameIndexes();

	/// Menu helper functions
	
	function SetCorrectFrameIndexes()
	{
		if(parent.frames.length == 2)
		{
			rootFrameIndex = -1; //Indicates that the root frame is not in use
			menuFrameIndex = 0;	 //The menu frame is now the root frame
			funcFrameIndex = 1; 
		}			
	}	
	
	function LoadMainPage(mainPageURL)
	{
		parent.frames[funcFrameIndex].location = mainPageURL;
		focus();
	}