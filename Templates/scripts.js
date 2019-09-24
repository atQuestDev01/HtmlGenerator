<script language="javascript">

    function GetContents()
    {
        alert(toJSONString(document.getElementById("formTable")));
    }

    function toJSONString(currentElem) {
        var obj = {};
        var elements = currentElem.querySelectorAll("input, select, textarea");
        for (var i = 0; i < elements.length; ++i) {
            var element = elements[i];
            var name = element.name;
            var value = element.value;
            var type = element.type;
            var checked = element.checked;

            if (name) {
                if (type == "radio" && checked) {        
                    obj[name] = value;
                }
                if (type != "radio") {
                    obj[name] = value;
                }
            }
        }

        return JSON.stringify(obj);
    }

</script>