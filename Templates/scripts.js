	var jsonObj = {};
    var jsonText = "";
	
	
	function SetContent() 
	{
		jsonObj = JSON.parse(document.getElementById("data").value);
		
		for (var key in jsonObj) 
		{
			var element = document.getElementById(key);
			
			switch (element.type) 
			{
				case "radio":
					var radioElems = document.getElementsByName(key);
					for (var i = 0; i < radioElems.length; i++)
					{
						if (radioElems[i].value == jsonObj[key]) 
						{
							radioElems[i].checked = true;
							break;
						}
					}									
					break;
				
				case "select-one":
					var options = element.options;					
					for (var i = 0; i < options.length; i++) 
					{
						if (options[i].value == jsonObj[key])
						{
							options[i].selected = true;
							break;
						}
					}				
					break;
					
				default:
					element.value = jsonObj.key;
			}
		}
	}
	
	function GetContent()
    {
        var currentElem = document.getElementById("formTable");
        var elements = currentElem.querySelectorAll("input, select, textarea");
        
		for (var i = 0; i < elements.length; ++i) {
            var element = elements[i];
            var name = element.name;
            var value = element.value;
            var type = element.type;
            var checked = element.checked;

            if (name) {
                if (type == "radio" && checked) {        
                    jsonObj[name] = value;
                }
                if (type != "radio") {
                    jsonObj[name] = value;
                }
            }
        }

        jsonText = JSON.stringify(jsonObj);
		
		document.getElementById("data").value = jsonText;
    }

