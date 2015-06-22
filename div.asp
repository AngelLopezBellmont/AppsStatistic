<%@  language="JScript" %>

<script language="javascript" runat="server"></script>


<%
    //var sql = GetConnectionBarForce();
    //strQuery = "select ProcessID, ProcessUID, Status from SlitteApp where ProcessID='129'"
    //rs = sql.execute(strQuery)

    //2014.11.28 Alopez i change the conexion so we dont need anymore the include from Ralph
    //var sql = GetConnectionBarForce();

    var cn = Server.CreateObject("ADODB.Connection");
    var ProvStr = "Provider=SQLNCLI10; Data Source=sql3; Initial Catalog=BarForce; Integrated Security=SSPI;"
    cn.Open(ProvStr);

    var sql = cn

    //Angel Lopez I delete this function because is one old that use inlclude
    //function getFieldValue_l(fldName) {
    //    return getFieldValue(rs, fldName);
    //}


    //2014.12.15 Alopez We control if today is monday, so we can show in Apps Gestern the Apps from today monday and friday instead from sunday
    var today = new Date();
    var today_getDay = today.getDay();

    var yesterday = new Date(today);
    var yesterday_getDay = 1;

    if (today_getDay==1){
        yesterday_getDay =3;
    }

    yesterday.setDate(today.getDate()- yesterday_getDay);


    var yesterday_getDay = yesterday.getDay();

    ////Response.Write( "today: " + today);
    ////Response.Write( " / today_getDay: " + today_getDay);

    //////Response.Write( "  // new Date().getTime():  " + new Date().getTime());
    ////Response.Write( "  // new Date().getDay():  " + yesterday.getDay());


    //Angel Lopez
    function add_null(my_Data) {
        if (my_Data < 10) {
            my_Data = "0" + my_Data;
        }

        else {
            my_Data = my_Data;
        }

        return my_Data;
    }


    function mynow() {
        var my_today = new Date();

        my_today = change_Format_Time(my_today);
        return my_today;
    }

    //Angel Lopez
    function change_Format_Time(myTime)
    {
        if (myTime == null || myTime == "")
            return "";

        var my_date = myTime.getDate();
        my_date = add_null(my_date);

        var my_month = myTime.getMonth();
        my_month = my_month + 1;
        my_month = add_null(my_month);

        //my_month_INT = parseInt(my_month);
        //my_month_INT = my_month_INT + 1;
        //my_month_INT = add_null(my_month_INT);

        var my_year = myTime.getFullYear();


        var my_hour = (myTime.getHours());
        my_hour = add_null(my_hour);

        var my_minute = myTime.getMinutes();
        my_minute = add_null(my_minute);

        var my_secunden = myTime.getSeconds();
        my_secunden = add_null(my_secunden);

        myTime = my_date + "-" + my_month + "-" + my_year + "  " + my_hour + ":" + my_minute + ":" + my_secunden;
        return myTime;
    }

    //Angel Lopez
    function change_Date_Format_Field(myField_Date)
    {
        if (myField_Date == null || myField_Date == "")
            return "";

        strQuery = "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY,"+ myField_Date +") AS varchar(2)), 2) + '.' +"
                    + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, "+ myField_Date +") AS varchar(2)), 2) + '.' + "
                    + "CAST(YEAR("+ myField_Date +") AS VARCHAR(4)) +'  '+ "
                    + "Convert(varchar(5), "+ myField_Date +", 108) "

        return strQuery;
    }



    //Angel Lopez
    function change_Format_Number_old (number)
    {
        if (number >= 1000 && number <10000 )
        {
            number =number+ '';
            var res1 = number.substr(0, 1)
            var res2 = number.substr(1, 3)

            number = res1 +"."+res2
        }

        if (number >=10000 && number<100000 )
        {
            number =number+ '';
            var res1 = number.substr(0, 2)
            var res2 = number.substr(2, 3)

            number = res1 +"."+res2
        }

        if (number >=100000 && number<1000000 )  
        {
            number =number+ '';
            var res1 = number.substr(0, 3)
            var res2 = number.substr(3, 3)

            number = res1 +"."+res2
        }

        if (number >=1000000 && number<10000000 )  
        {
            number =number+ '';
            var res1 = number.substr(0, 1)
            var res2 = number.substr(1, 3)
            var res3 = number.substr(4, 3)

            number = res1 +"."+res2+"."+res3
        }

        if (number >=10000000 && number<100000000 )  
        {
            number =number+ '';
            var res1 = number.substr(0, 2)
            var res2 = number.substr(2, 3)
            var res3 = number.substr(5, 3)

            number = res1 +"."+res2+"."+res3
        }

        if (number >=100000000 && number<1000000000 )  
        {
            number =number+ '';
            var res1 = number.substr(0, 3)
            var res2 = number.substr(3, 3)
            var res3 = number.substr(6, 3)

            number = res1 +"."+res2+"."+res3
        }

        return number;

    }

    //2014.12.15 Alopez we write a . in thousand, millon etc.
    function change_Format_Number (number)
    {
       
        number = number+"";
        
        n_length = number.length;
        var rem = n_length % 3;
       
        var  number_pre = number.substring(0,rem);
        var number_post = number.substring(rem,n_length );
        var number_post_b="";

        for (i = 0; i < number_post.length ; i++) 
        { 
            if ( (i%3) ==0){ 

                j = i-3;
                number_post_b = number_post_b + "." + number_post.substr(j, 3);
                //Response.Write ("i;" + i + "  j:" + j + number_post_b);
                //Response.Write ("<br/>");
            }
        }


        if (number_pre.length == 0)
        {
            number_post_b = number_post_b.substring(1,number_post_b.length)
            number= number_post_b;
        }

        number = number_pre + number_post_b;

        return  number;
        
    }


    //2014.12.12 Alopez
    function secondsToString(seconds)
    {
        var numyears = Math.floor(seconds / 31536000);
        var numdays = Math.floor((seconds % 31536000) / 86400); 
        var numhours = Math.floor(((seconds % 31536000) % 86400) / 3600);
        var numminutes = Math.floor((((seconds % 31536000) % 86400) % 3600) / 60);
        var numseconds = (((seconds % 31536000) % 86400) % 3600) % 60;

        var myString_full_Fecha;
        var myColor = "green";


        if ( numseconds ==1) { var string_seconds = "sec";}
        else{var string_seconds = "secs";}


        if ( numhours ==1) { var string_hour = "hour";}
        else{var string_hour = "hours";}


        if ( numdays ==1) { var string_days = "day"; }
        else 
        {
            var string_days = "days";
            if (numdays> 1) { myColor = "red"}
        }


        if ((numyears == 0) && (numdays == 0) && (numhours == 0) && (numminutes == 0) && (numseconds != 0)) { 
            var myString_full_Fecha = numseconds + " seconds";}


        if ((numyears == 0) && (numdays == 0) && (numhours == 0) && (numminutes != 0) && (numseconds != 0))  {  
            myString_full_Fecha = numminutes + " min  " + numseconds + " sec" ;}

        if ((numyears == 0) && (numdays == 0) && (numhours != 0) && (numminutes != 0) && (numseconds != 0))  {  

         
            myString_full_Fecha = numhours + " "+ string_hour + " " + numminutes + " min ";
       
             }

        if ((numyears == 0) && (numdays != 0) && (numhours != 0) && (numminutes != 0) && (numseconds != 0)) {  
            //myString_full_Fecha =  numdays + " days " + numhours + " hours " + numminutes + " minutes " + numseconds + " seconds";

            myString_full_Fecha =  numdays + " "+ string_days +" " + numhours + " " + string_hour;
        }

        if ((numyears != 0) && (numdays != 0) && (numhours != 0) && (numminutes != 0) && (numseconds != 0)) {  

            myString_full_Fecha = numyears + " years  " +  numdays + " days ";
        }

        var myArray = [myString_full_Fecha, myColor];

        return myArray;
    }



    //2014.12.12 Angel Lopez
    function get_TimeLeft_to_now (myValue_Datum_Formatiert)
    {

      
        //var dt = new Date(2000, 10, 1, 18, 45, 45 );      // si funciona  [Wed Nov 1 18:45:45 UTC+0100 2000]
        //var dt_getTime = dt.getTime();                    // si funciona  [973100745000]

        // var momentDate =  Date('2025-09-15 09:00');       // si  funciona me da la fecha Now() la fecha                             ej:Fri Dec 12 12:17:00 2014
                
        var time_1A = new Date();                          // Si funciona me devuelve now()                                          [Fri Dec 12 11:43:39 UTC+0100 2014]
        //var time_1B_2 = myValue_Datum_Formatiert;          // Si funciona                                                            [12.12.2014 12:36]

        

        //var time_1C = parseDate('31.05.2010', 'dd.mm.yyyy');      // NO funciona
        //var time_1D = parseDate(myValue_Datum_OhneFormat);        // NO funciona
                

        //var n = myValue_Datum_Formatiert.length;     // 17  [12.12.2014 12:36]
        //var sub_string17 = myValue_Datum_Formatiert.substring(16,17);
        //var sub_string11 = myValue_Datum_Formatiert.substring(11,12);  // null value
        //var sub_string10 = myValue_Datum_Formatiert.substring(10,11);  // null value
              
        var myValue_Datum_Formatiert_B = myValue_Datum_Formatiert.substring(0,10)+"."+ myValue_Datum_Formatiert.substring(12,17);   //12.12.2014.12:47 
        var myValueToSplit = myValue_Datum_Formatiert_B.replace(':', '.');  // 12.12.2014.12.47

        var myArr_split = myValueToSplit.split("."); 
        var myDate_fromString_to_DateType = new Date(myArr_split[2], myArr_split[1] - 1, myArr_split[0], myArr_split[3], myArr_split[4]);

        var time_2 = new Date().getTime();                 // Si funciona me devuelve los milisegundos desde 1 Jan 1970, midnight    "1418380921461"
        var time_2B = myDate_fromString_to_DateType.getTime();                    //No funciona no estoy seguro que myValue_Datum_OhneFormat sea una date value

  
        //var time_3 = heute.getTime();                             // no funciona
        //var time_4 =  Date(myValue_Datum_Formatiert);             // no funciona me devuleve la fecha de hoy no la de 

        //var time_5 = (time_2) - (time_2B);                                     // 4026980 
        var diff = Math.abs(time_1A - myDate_fromString_to_DateType);          // 4026980
        var mySeconds =  Math.round((diff)/1000);

        return mySeconds;
    }


    var myColorGreen = "#73bf26";
    var myColorRed =  "#d13426";

%>

<html lang="en-US">

<head>
    <title></title>

    <style>
        .miDemo {
            /*position:absolute;*/
            float: left;
            top: 50%;
            /*margin-top:-50px;*/
            left: 50%;
            /*margin-left: -50px ;*/
            margin-left: auto;
            margin-right: auto;
            width: 100px;
            height: 100px;
            /*margin:50px;*/
            /*padding: 15px;*/
            border: 1px solid black;
            background-color: #ff0000;
        }

        .big_box_up {
            /*background-color: #5798d9; border: 1px solid black; position: absolute; width: 1250; height: 550px; top: 50%; left: 50%; margin-top: -275px; margin-left: -625px; text-align: center;*/
            background-color: #5798d9;
            position: absolute;
            width: 1220px;
            height: 550px;
            top: 50%;
            left: 50%;
            margin-top: -280px;
            margin-left: -610px;
            text-align: center;
            border: 2px solid #5798d9;
            background-image: url('allinfos_background.png');
        }

        /*.big_box_down {
            background-color: #5798d9;
            border: 1px solid black;
            position: absolute;
            width: 1160px;
            height: 550px;
            top: 50%;
            left: 50%;
            margin-top: -275px;
            margin-left: -580px;
            text-align: center;
            background-color: #808080;
            position: absolute;
            width: 1240px;
            height: 550px;
            top: 100%;
            left: 50%;
            margin-top: -275px;
            margin-left: -625px;
            text-align: center;
            border: 1px solid black;
        }*/

        .box {
            font-size: initial;
            float: left;
            border: 1px solid #ccc;
            background-color: #FFFFFF;
            width: 250px;
            height: 425px;
            margin-left: 20px;
            margin-top: 25px;
            /*color: #808080;
            font-family: arial;
            /*font-size: 11pt;*/ */;
            /*font-weight: bold;*/
            /*margin: 80px;*/
            /*-moz-box-shadow: 0px 0px 8px rgba(68,68,68,0.4);
            -webkit-box-shadow: 0px 0px 8px rgba(68,68,68,0.4);
            box-shadow: 0px 0px 8px rgba(68,68,68,0.4);*/
        }

        .Kopf {
            color: #F8F8FF;
            font-family: arial;
            font-size: 22pt;
            font-weight: bold;
            margin-top: 25px;
        }

        .Foot {
            color: #8c8c8c;
            font-family: arial;
            font-size: 12pt;
            margin-bottom: 20px;
            margin-top: 460px;
            /*border: 1px solid black;*/
        }


        .Titel_Box {
            color: #808080;
            font-family: arial;
            /*font-size: initial;*/
            font-size: 14pt;
            margin-top: 14px;
            margin-bottom: 14px;
        }


        .text_oben_grey {
            color: #8c8c8c;
            font-family: arial;
            /*font-size: initial;*/
            font-size: 11pt;
            font-weight: bold;
            margin-bottom: 2px;
        }

        .letter_to_left {
            float: left;
            color: #8c8c8c;
            font-family: arial;
            /*font-size: initial;*/
            font-size: 10pt;
            /*font-size-adjust: 0.58;*/
            text-align: left;
            margin-left: 10px;
            /*border: 1px solid;*/
        }

        .letter_to_right {
            /*color: #1E90FF;*/
            color: #3687d9;        
            font-family: arial;
            /*font-size: initial;*/
            font-size: 10pt;
            /*font-size-adjust: 0.58;*/
            text-align: right;
            float: right;
            /*margin-left:10px;*/
            margin-right: 11px;
            /*border: 1px solid;*/
        }

        .letter_to_right_red {
            /*color: #1E90FF;*/
            color: #d13426;;        
            font-family: arial;
            /*font-size: initial;*/
            font-size: 10pt;
            /*font-size-adjust: 0.58;*/
            text-align: right;
            float: right;
            /*margin-left:10px;*/
            margin-right: 11px;
            /*border: 1px solid;*/
        }

        .letter_green_date {
            color: #73bf26;
         
            font-family: arial;
            /*font-size: initial;*/
            font-size: 11pt;
            font-weight: bold;
            margin-bottom: 7px;
        }
        .letter_red_date {
            color: #d13426;
            font-family: arial;
            /*font-size: initial;*/
            font-size: 11pt;
            font-weight: bold;
            margin-bottom: 7px;
        }


        .letter_green_number {
            color: #73bf26;
      
            font-family: arial;
            /*font-size: initial;*/
            font-size: 15pt;
            font-weight: bold;
            margin-bottom: 7px;
        }

        .letter_red_number {
            color: #d13426;

            font-family: arial;
            /*font-size: initial;*/
            font-size: 15pt;
            font-weight: bold;
            margin-bottom: 7px;
        }

        .hline {
            width: 100%;
            height: 1px;
            background: #fff;
        }

        .shadow1 {
            margin: 80px;
            background-color: rgb(68,68,68);
            -moz-box-shadow: 0px 0px 8px rgba(68,68,68,0.8);
            -webkit-box-shadow: 0px 0px 8px rgba(68,68,68,0.8);
            box-shadow: 0px 0px 8px rgba(68,68,68,0.8);
        }

        hr {
            width: 230px;
            height: 1px;
            margin-right: 10px;
            margin-left: 10px;
            border-style: solid;
            border-width: 1px 0 0;
            border-color: #ccc;
        }
    </style>

</head>


<body>


    <div class="big_box_up">
        <!-- <div class="big_box_down">-->


        <div class="Kopf">
            <img src="allinfos_title.png" />
        </div>

        <div class="box" style="margin-left: 80px">


            <div class="Titel_Box">
                <img src="apple.jpg" />
            </div>


            <div class="text_oben_grey">Letzte App hochgeladen </div>
            <%

                //strQuery_ohne_Format = "SELECT TOP 1  iOS_EndTime FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"
                // strQuery = "SELECT TOP 1 convert(VARCHAR(16), iOS_EndTime, 101) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"
                // strQuery = "SELECT TOP 1 convert(VARCHAR(30), iOS_EndTime, 10) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"
                // strQuery = "SELECT TOP 1 format( iOS_EndTime, 'dd-mm-yyyy HH:m:ss') FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"  //Let the date like this 24-19-2014 05:19:10
                // strQuery = "SELECT TOP 1 convert(VARCHAR, iOS_EndTime, 113) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"
                // strQuery = "SELECT TOP 1 convert(VARCHAR, iOS_EndTime, 20) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"                                      // query good 
                // strQuery = "SELECT dbo.TobitFormatDate (iOS_EndTime, '#dd#.#mm#.#yyyy# #hh#:#mi#', 'eng') FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"       // query good
                // strQuery = "SELECT TOP 1 convert(date, iOS_EndTime) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"
                // strQuery = "select CONVERT(VARCHAR(10), GETDATE(), 4) + ' '  + convert(VARCHAR(8), GETDATE(), 8)"
                // strQuery = "SELECT CONVERT(VARCHAR(19), GETDATE(), 110) AS [MM-DD-YYYY] " 
                //  strQuery = "SELECT CAST(DAY(GETDATE()) AS VARCHAR(2)) + '-' + CAST(MONTH(GETDATE()) AS VARCHAR(2))+ '-' + CAST(YEAR(GETDATE()) AS VARCHAR(4)) + '  ' + Convert(varchar(5), GetDate(), 108)"  //very good Angel du bist Super :)

                //strQuery = "SELECT CAST(DAY(iOS_EndTime) AS VARCHAR(2)) + '-' + CAST(MONTH(iOS_EndTime) AS VARCHAR(2))+ '-' + CAST(YEAR(iOS_EndTime) AS VARCHAR(4)) +'  '+ Convert(varchar(5), iOS_EndTime, 108)"
                //+ "FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC"

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, iOS_EndTime) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, iOS_EndTime) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(iOS_EndTime) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), iOS_EndTime, 108) "
                //            + "FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC";

                var sub_query = change_Date_Format_Field("iOS_EndTime");
                strQuery_Datum_Formiert = "SELECT TOP 1" + sub_query + "FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL ORDER BY iOS_EndTime DESC";

  

                var rs = sql.execute(strQuery_Datum_Formiert)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    //Response.Write(change_Format_Time(myValue) + "<br/>");
                    //Response.Write(change_Format_Time(getFieldValue_l("Android_EndTime")) + "<br/>");  //works ok, with normal query

                    rs.movenext();
                }


                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);
                
                if (str_full_mySeconds[1] == "green") {  Response.Write("<div class=\"letter_green_date\">"); }
                if (str_full_mySeconds[1] == "red") {  Response.Write("<div class=\"letter_red_date\">"); }

                //Response.Write("<div class=\"letter_green_date\">");
                //Response.Write(myValue_Datum_Formatiert + " ---- " +   myValueToSplit);
                Response.Write (str_full_mySeconds[0]);
                //Response.Write("</div><br/>");
                Response.Write("</div>");



                
            %>


            <div class="text_oben_grey">Gestern/Heute hochgeladen   </div>

            <%
                // strQuery = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL AND convert(VARCHAR, iOS_EndTime, 20) = Convert(VARCHAR, GETDATE(),20)"

                //strQuery_today = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL AND CONVERT(date,iOS_EndTime) = Convert(date, GETDATE())"    //good query

                // yesterday.setDate(today.getDate()- yesterday_getDay);

                strQuery_today = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL AND CONVERT(date,iOS_EndTime) = Convert(date, GETDATE())";

                
                
                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL AND CONVERT(date,iOS_EndTime) = Convert(date, GETDATE()-1)";
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE iOS_EndTime IS NOT NULL AND CONVERT(date,iOS_EndTime) >= Convert(date, GETDATE()-3) AND CONVERT(date,iOS_EndTime) < Convert(date, GETDATE())";}
               
                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");
            %>


            <div class="text_oben_grey">Gestern/Heute freigegeben   </div>
            <%

                strQuery_today = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnIOSDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnIOSDate) = Convert(date, GETDATE())";
                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnIOSDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnIOSDate) = Convert(date, GETDATE()-1)";
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnIOSDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnIOSDate) >= Convert(date, GETDATE()-3) AND CONVERT(date,FirstDownloadBtnIOSDate) < Convert(date, GETDATE())";}

                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");
            %>

            <div class="letter_to_left">LastDownloadButton </div>
            <!-- <br />-->
            <%
                // strQuery = "SELECT TOP 1 CAST(DAY(FirstDownloadBtnIOSDate) AS VARCHAR(2)) + '-' + CAST(MONTH(FirstDownloadBtnIOSDate) AS VARCHAR(2))+ '-' + CAST(YEAR(FirstDownloadBtnIOSDate) AS VARCHAR(4)) +'  '+ Convert(varchar(5), FirstDownloadBtnIOSDate, 108) " +
                //    "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnIOSDate IS NOT NULL ORDER BY FirstDownloadBtnIOSDate DESC";

                // strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, FirstDownloadBtnIOSDate) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, FirstDownloadBtnIOSDate) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(FirstDownloadBtnIOSDate) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), FirstDownloadBtnIOSDate, 108) "
                //            + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnIOSDate IS NOT NULL ORDER BY FirstDownloadBtnIOSDate DESC";


                var sub_query = change_Date_Format_Field("FirstDownloadBtnIOSDate");
                strQuery = "SELECT TOP 1" + sub_query + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnIOSDate IS NOT NULL ORDER BY FirstDownloadBtnIOSDate DESC";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);  //2014.12.15 Alopez now str_full_mySeconds is one array because I chnage the function secondsToString an now return is one array
              
                if (str_full_mySeconds[1] == "green") { Response.Write("<div class=\"letter_to_right\">"); }
                if (str_full_mySeconds[1] == "red")   { Response.Write("<div class=\"letter_to_right_red\">");}

                //Response.Write("<div class=\"letter_to_right\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");

            %>

            <hr />
            <div class="letter_to_left">Letztes Zertifikat </div>
            <!-- <br />-->
            <%
                //strQuery = "SELECT TOP 1   CAST(DAY(CertificateReadyDate) AS VARCHAR(2)) + '-' + CAST(MONTH(CertificateReadyDate) AS VARCHAR(2))+ '-' + CAST(YEAR(CertificateReadyDate) AS VARCHAR(4)) +'  '+ Convert(varchar(5), CertificateReadyDate, 108)  " +
                //    "FROM SlitteApp (NOLOCK) WHERE CertificateReadyDate IS NOT NULL ORDER BY CertificateReadyDate DESC";

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, CertificateReadyDate) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, CertificateReadyDate) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(CertificateReadyDate) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), CertificateReadyDate, 108) "
                //            + "FROM SlitteApp (NOLOCK) WHERE CertificateReadyDate IS NOT NULL ORDER BY CertificateReadyDate DESC";

                var sub_query = change_Date_Format_Field("CertificateReadyDate");
                strQuery = "SELECT TOP 1" + sub_query + "FROM SlitteApp (NOLOCK) WHERE CertificateReadyDate IS NOT NULL ORDER BY CertificateReadyDate DESC";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") { Response.Write("<div class=\"letter_to_right\">"); }
                if (str_full_mySeconds[1] == "red")   { Response.Write("<div class=\"letter_to_right_red\">");}

                //Response.Write("<div class=\"letter_to_right\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");


            %>

            <div class="letter_to_left">Zertifikate heute </div>
            <!--<br />-->
            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE CertificateReadyDate IS NOT NULL AND CONVERT(date,CertificateReadyDate) = Convert(date, GETDATE())";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(change_Format_Number(rs(0)));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>

            <hr />
            <div class="letter_to_left">WaitingForReview heute </div>
            <!-- <br />-->
            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstWaitingForReleaseDate IS NOT NULL AND CONVERT(date,FirstWaitingForReleaseDate) = Convert(date, GETDATE())";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(rs(0));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>
            <div class="letter_to_left">WaitingForReview gestern </div>
            <!--  <br />-->
            <%

                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstWaitingForReleaseDate IS NOT NULL AND CONVERT(date,FirstWaitingForReleaseDate) = Convert(date, DATEADD(dd, -1, GETDATE()))";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(rs(0));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>
            <hr />
            <div class="letter_to_left">Zuletzt modifiziert  </div>
            <!-- <br />-->
            <%

                //strQuery = "SELECT TOP 1 CAST(DAY(ITC_LastModified1) AS VARCHAR(2)) + '-' + CAST(MONTH(ITC_LastModified1) AS VARCHAR(2))+ '-' + CAST(YEAR(ITC_LastModified1) AS VARCHAR(4)) +'  '+ Convert(varchar(5), ITC_LastModified1, 108)"
                //+ "FROM SlitteApp (NOLOCK) WHERE ITC_LastModified1 IS NOT NULL ORDER BY ITC_LastModified1 DESC";

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, ITC_LastModified1) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, ITC_LastModified1) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(ITC_LastModified1) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), ITC_LastModified1, 108) "
                //            + "FROM SlitteApp (NOLOCK) WHERE ITC_LastModified1 IS NOT NULL ORDER BY ITC_LastModified1 DESC";

                var sub_query = change_Date_Format_Field("ITC_LastModified1");
                strQuery = "SELECT TOP 1" + sub_query + "FROM SlitteApp (NOLOCK) WHERE ITC_LastModified1 IS NOT NULL ORDER BY ITC_LastModified1 DESC";


                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") { Response.Write("<div class=\"letter_to_right\">"); }
                if (str_full_mySeconds[1] == "red")   { Response.Write("<div class=\"letter_to_right_red\">");}

                //Response.Write("<div class=\"letter_to_right\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");

            %>

            <div class="letter_to_left">Anzahl modifiziert </div>
            <!-- <br />-->
            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE ITC_LastModified1 IS NOT NULL AND CONVERT(date,ITC_LastModified1) = Convert(date, GETDATE())";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(rs(0));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>
        </div>


        <div class="box">

            <div class="Titel_Box">
                <img src="android.jpg" />
            </div>


            <div class="text_oben_grey">Letzte App hochgeladen </div>

            <%
                //strQuery = "SELECT TOP 1  CAST(DAY(Android_EndTime) AS VARCHAR(2)) + '-' + CAST(MONTH(Android_EndTime) AS VARCHAR(2))+ '-' + CAST(YEAR(Android_EndTime) AS VARCHAR(4)) +'  '+ Convert(varchar(5), Android_EndTime, 108) "
                //   + "FROM ChaynsProductionStatus (NOLOCK) WHERE Android_EndTime IS NOT NULL ORDER BY Android_EndTime DESC"

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, Android_EndTime) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, Android_EndTime) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(Android_EndTime) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), Android_EndTime, 108) "
                //            + "FROM ChaynsProductionStatus (NOLOCK) WHERE Android_EndTime IS NOT NULL ORDER BY Android_EndTime DESC";

                var sub_query = change_Date_Format_Field("Android_EndTime");
                strQuery = "SELECT TOP 1" + sub_query + "FROM ChaynsProductionStatus (NOLOCK) WHERE Android_EndTime IS NOT NULL ORDER BY Android_EndTime DESC";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") {  Response.Write("<div class=\"letter_green_date\">"); }
                if (str_full_mySeconds[1] == "red") {  Response.Write("<div class=\"letter_red_date\">"); }

                //Response.Write("<div class=\"letter_green_date\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");

            %>

            <div class="text_oben_grey">Gestern/Heute hochgeladen </div>
            <%
                strQuery_today = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE Android_EndTime IS NOT NULL AND CONVERT(date,Android_EndTime) = Convert(date, GETDATE())"
               



                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE Android_EndTime IS NOT NULL AND CONVERT(date,Android_EndTime) = Convert(date, GETDATE()-1)"
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE Android_EndTime IS NOT NULL AND CONVERT(date,Android_EndTime) >= Convert(date, GETDATE()-3) AND CONVERT(date,Android_EndTime) < Convert(date, GETDATE())";}

                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");
            %>

            <div class="text_oben_grey">Gerstern/Heute freigegeben   </div>
            <%
                strQuery_today = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAndroidDate) = Convert(date, GETDATE())";
                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAndroidDate) = Convert(date, GETDATE()-1)";
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAndroidDate) >= Convert(date, GETDATE()-3) AND CONVERT(date,FirstDownloadBtnAndroidDate) < Convert(date, GETDATE())";}
                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }


                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");
            %>


            <div class="letter_to_left">LastDownloadButton </div>
            <%
                //strQuery = "SELECT TOP 1 convert(VARCHAR, FirstDownloadBtnAndroidDate, 20)  FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL ORDER BY FirstDownloadBtnAndroidDate DESC";
                //strQuery = "SELECT TOP 1 CAST(DAY(FirstDownloadBtnAndroidDate) AS VARCHAR(2)) + '-' + CAST(MONTH(FirstDownloadBtnAndroidDate) AS VARCHAR(2))+ '-' + CAST(YEAR(FirstDownloadBtnAndroidDate) AS VARCHAR(4)) +'  '+ Convert(varchar(5), FirstDownloadBtnAndroidDate, 108) "
                //+ "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL ORDER BY FirstDownloadBtnAndroidDate DESC";

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, FirstDownloadBtnAndroidDate) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, FirstDownloadBtnAndroidDate) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(FirstDownloadBtnAndroidDate) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), FirstDownloadBtnAndroidDate, 108) "
                //            + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL ORDER BY FirstDownloadBtnAndroidDate DESC";


                var sub_query = change_Date_Format_Field("FirstDownloadBtnAndroidDate");
                strQuery = "SELECT TOP 1" + sub_query + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAndroidDate IS NOT NULL ORDER BY FirstDownloadBtnAndroidDate DESC";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") { Response.Write("<div class=\"letter_to_right\">"); }
                if (str_full_mySeconds[1] == "red")   { Response.Write("<div class=\"letter_to_right_red\">");}

                //Response.Write("<div class=\"letter_to_right\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");

            %>

            <hr />
            <div class="letter_to_left">Total APKs Cloud </div>
            <%

                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAPKDate IS NOT NULL";
                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    //Response.Write(rs(0));
                    //Response.Write(change_Format_Number(123456789));

                    Response.Write(change_Format_Number(rs(0)));
                   
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>

            <div class="letter_to_left">APKs Cloud Heute </div>

            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAPKDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAPKDate) = Convert(date, GETDATE())";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(change_Format_Number(rs(0)));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>

            <hr />
            <div class="letter_to_left">DownloadButtonAndroid=1 </div>

            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE DownloadBtnAndroid =1";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(change_Format_Number(rs(0)));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>
            <div class="letter_to_left">DownloadButtonAndroid=1 & </div>
            <br />
            <div class="letter_to_left">Android Version < 4097 </div>
            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE DownloadBtnAndroid =1 and AndroidVersion < 4097";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(change_Format_Number(rs(0)));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }

            %>
        </div>


        <div class="box">

            <div class="Titel_Box">
                <img src="amazon.jpg" />
            </div>

            <div class="text_oben_grey">Letzte App hochgeladen  </div>

            <%
                //strQuery = "SELECT TOP 1 CAST(DAY(Amazon_EndTime) AS VARCHAR(2)) + '-' + CAST(MONTH(Amazon_EndTime) AS VARCHAR(2))+ '-' + CAST(YEAR(Amazon_EndTime) AS VARCHAR(4)) +'  '+ Convert(varchar(5), Amazon_EndTime, 108)"
                //+ "FROM ChaynsProductionStatus (NOLOCK) WHERE Amazon_EndTime IS NOT NULL ORDER BY Amazon_EndTime DESC"

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, Amazon_EndTime) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, Amazon_EndTime) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(Amazon_EndTime) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), Amazon_EndTime, 108) "
                //            + "FROM ChaynsProductionStatus (NOLOCK) WHERE Amazon_EndTime IS NOT NULL ORDER BY Amazon_EndTime DESC";

                var sub_query = change_Date_Format_Field("Amazon_EndTime");
                strQuery = "SELECT TOP 1" + sub_query + "FROM ChaynsProductionStatus (NOLOCK) WHERE Amazon_EndTime IS NOT NULL ORDER BY Amazon_EndTime DESC";


                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") {  Response.Write("<div class=\"letter_green_date\">"); }
                if (str_full_mySeconds[1] == "red") {  Response.Write("<div class=\"letter_red_date\">"); }

                //Response.Write("<div class=\"letter_green_date\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");
                
            %>


            <div class="text_oben_grey">Gestern/Heute hochgeladen  </div>

            <%
                strQuery_today = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE Amazon_EndTime IS NOT NULL AND CONVERT(date,Amazon_EndTime) = Convert(date, GETDATE())"
               

                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE Amazon_EndTime IS NOT NULL AND CONVERT(date,Amazon_EndTime) = Convert(date, GETDATE()-1)"
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE Amazon_EndTime IS NOT NULL AND CONVERT(date,Amazon_EndTime) >= Convert(date, GETDATE()-3) AND CONVERT(date,Amazon_EndTime) < Convert(date, GETDATE())";}

                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");
            %>

            <div class="text_oben_grey">Gerstern/Heute freigegeben   </div>

            <%
                strQuery_today = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAmazonDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAmazonDate) = Convert(date, GETDATE())";

                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAmazonDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAmazonDate) = Convert(date, GETDATE())";
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAmazonDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnAmazonDate) >= Convert(date, GETDATE()-3) AND CONVERT(date,FirstDownloadBtnAmazonDate) < Convert(date, GETDATE())";}
                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");

            %>


            <div class="letter_to_left">LastDownloadButton </div>

            <%
                //strQuery = "SELECT TOP 1 CAST(DAY(FirstDownloadBtnAmazonDate) AS VARCHAR(2)) + '-' + CAST(MONTH(FirstDownloadBtnAmazonDate) AS VARCHAR(2))+ '-' + CAST(YEAR(FirstDownloadBtnAmazonDate) AS VARCHAR(4)) +'  '+ Convert(varchar(5), FirstDownloadBtnAmazonDate, 108) "
                //+ "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAmazonDate IS NOT NULL ORDER BY FirstDownloadBtnAmazonDate DESC";

                //strQuery = "select RIGHT( REPLICATE('0', 2) + CAST(DATEPART(DAY, '2012-12-08') AS varchar(2)), 2)"

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, FirstDownloadBtnAmazonDate) AS varchar(2)), 2) + '-' +"
                //+ "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, FirstDownloadBtnAmazonDate) AS varchar(2)), 2) + '-' + "
                //+ "CAST(YEAR(FirstDownloadBtnAmazonDate) AS VARCHAR(4)) +'  '+ "
                //+ "Convert(varchar(5), FirstDownloadBtnAmazonDate, 108) "
                //+ "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAmazonDate IS NOT NULL ORDER BY FirstDownloadBtnAmazonDate DESC";

                var sub_query = change_Date_Format_Field("FirstDownloadBtnAmazonDate");
                strQuery = "SELECT TOP 1" + sub_query + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnAmazonDate IS NOT NULL ORDER BY FirstDownloadBtnAmazonDate DESC";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") { Response.Write("<div class=\"letter_to_right\">"); }
                if (str_full_mySeconds[1] == "red")   { Response.Write("<div class=\"letter_to_right_red\">");}

                //Response.Write("<div class=\"letter_to_right\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");
            %>
        </div>


        <div class="box">

            <div class="Titel_Box">
                <img src="windows.jpg" />
            </div>

            <div class="text_oben_grey">Letzte App hochgeladen </div>

            <%
                //strQuery = "SELECT TOP 1 convert(VARCHAR, WP8_EndTime, 20) FROM ChaynsProductionStatus (NOLOCK) WHERE WP8_EndTime IS NOT NULL ORDER BY WP8_EndTime DESC"

                strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, WP8_EndTime) AS varchar(2)), 2) + '-' +"
                            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, WP8_EndTime) AS varchar(2)), 2) + '-' + "
                            + "CAST(YEAR(WP8_EndTime) AS VARCHAR(4)) +'  '+ "
                            + "Convert(varchar(5), WP8_EndTime, 108) "
                            + "FROM ChaynsProductionStatus (NOLOCK) WHERE WP8_EndTime IS NOT NULL ORDER BY WP8_EndTime DESC";

                var sub_query = change_Date_Format_Field("WP8_EndTime");
                strQuery = "SELECT TOP 1" + sub_query + "FROM ChaynsProductionStatus (NOLOCK) WHERE WP8_EndTime IS NOT NULL ORDER BY WP8_EndTime DESC";


                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                if (str_full_mySeconds[1] == "green") {  Response.Write("<div class=\"letter_green_date\">"); }
                if (str_full_mySeconds[1] == "red") {  Response.Write("<div class=\"letter_red_date\">"); }

                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");
            %>

            <div class="text_oben_grey">Gestern/Heute hochgeladen </div>

            <%
                strQuery_today = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE WP8_EndTime  IS NOT NULL AND CONVERT(date,WP8_EndTime) = Convert(date, GETDATE())"
                


                rs = sql.execute(strQuery_today)
                if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE WP8_EndTime  IS NOT NULL AND CONVERT(date,WP8_EndTime) = Convert(date, GETDATE()-1)"
                if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM ChaynsProductionStatus (NOLOCK) WHERE WP8_EndTime IS NOT NULL AND CONVERT(date,WP8_EndTime) >= Convert(date, GETDATE()-3) AND CONVERT(date,WP8_EndTime) < Convert(date, GETDATE())";}

                rs = sql.execute(strQuery_yesterday)
                if (!rs.EOF) {
                    var myValue_yesterday= rs(0) + "";
                    rs.movenext();
                }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");

            %>

            <div class="text_oben_grey">Gerstern/Heute freigegeben  </div>

            <%
                 strQuery_today = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnWPDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnWPDate) = Convert(date, GETDATE())";

                 rs = sql.execute(strQuery_today)
                 if (!rs.EOF) {
                    var myValue_Today= rs(0) + "";
                    rs.movenext();
                }

                 strQuery_yesterday = "                       SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnWPDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnWPDate) = Convert(date, GETDATE()-1)";
                 if( today_getDay ==1){ strQuery_yesterday = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnWPDate IS NOT NULL AND CONVERT(date,FirstDownloadBtnWPDate) >= Convert(date, GETDATE()-3) AND CONVERT(date,FirstDownloadBtnWPDate) < Convert(date, GETDATE())";}
                 rs = sql.execute(strQuery_yesterday)
                 if (!rs.EOF) {
                     var myValue_yesterday= rs(0) + "";
                     rs.movenext();
                 }

                Response.Write("<div class=\"letter_green_number\">");
                Response.Write(myValue_yesterday +" / " + myValue_Today);
                Response.Write("</div>");
            %>


            <div class="letter_to_left">LastDownloadButton </div>

            <%

               // strQuery = "SELECT TOP 1 convert(VARCHAR, FirstDownloadBtnWPDate, 20)  FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnWPDate IS NOT NULL ORDER BY FirstDownloadBtnWPDate DESC";

                //strQuery = "SELECT TOP 1 RIGHT(REPLICATE('0', 2) + CAST(DATEPART(DAY, FirstDownloadBtnWPDate) AS varchar(2)), 2) + '-' +"
                //            + "RIGHT(REPLICATE('0', 2) + CAST(DATEPART(MONTH, FirstDownloadBtnWPDate) AS varchar(2)), 2) + '-' + "
                //            + "CAST(YEAR(FirstDownloadBtnWPDate) AS VARCHAR(4)) +'  '+ "
                //            + "Convert(varchar(5), FirstDownloadBtnWPDate, 108) "
                //            + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnWPDate IS NOT NULL ORDER BY FirstDownloadBtnWPDate DESC";

                var sub_query = change_Date_Format_Field("FirstDownloadBtnWPDate");
                strQuery = "SELECT TOP 1"+ sub_query + "FROM SlitteApp (NOLOCK) WHERE FirstDownloadBtnWPDate IS NOT NULL ORDER BY FirstDownloadBtnWPDate DESC";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    var myValue_Datum_Formatiert = rs(0) + "";
                    rs.movenext();
                }

                var mySeconds = get_TimeLeft_to_now (myValue_Datum_Formatiert);
                var str_full_mySeconds = secondsToString(mySeconds);

                
                if (str_full_mySeconds[1] == "green") { Response.Write("<div class=\"letter_to_right\">"); }
                if (str_full_mySeconds[1] == "red")   { Response.Write("<div class=\"letter_to_right_red\">");}

                //Response.Write("<div class=\"letter_to_right\">");
                Response.Write(str_full_mySeconds[0]);
                Response.Write("</div>");
            %>

            <hr />
            <div class="letter_to_left">Total XAPs Cloud </div>
            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE DownloadBtnXAP =1";
                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(change_Format_Number(rs(0)));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }
            %>

            <div class="letter_to_left">XAPs Cloud heute  </div>
            <%
                strQuery = "SELECT COUNT(SiteID) FROM SlitteApp (NOLOCK) WHERE DownloadBtnXAP =1  AND CONVERT(date,FirstDownloadBtnXAPDate) = Convert(date, GETDATE())";

                rs = sql.execute(strQuery)
                if (!rs.EOF) {
                    Response.Write("<div class=\"letter_to_right\">");
                    Response.Write(rs(0));
                    Response.Write("</div> <br/>");
                    rs.movenext();
                }
            %>
        </div>




        <div class="Foot">

            <!--        <div class="letter_to_left" style="float: left;">Inproduction: </div>-->
            Inproduction:
            <%
                var my_Inproduction;
                strQuery = "Select count(*)  from SlitteApp where Status ='INPRODUCTION'";
                rs = sql.execute(strQuery)

                if (!rs.EOF) {
                    my_Inproduction = rs(0);
            %>
            <span style="font-weight: bold;"><% =rs(0) %></span>
            <%
                //Response.Write(rs(0) + "&nbsp; &nbsp;");
                Response.Write("&nbsp; &nbsp; &nbsp;");
                rs.movenext();
                }


            %>

            <!-- <div class="letter_to_left" style="float: left;"> WaitingForRealease: </div>-->
            WaitingForRealease:
            <%
                strQuery = "Select count(*)  from SlitteApp where Status ='WAITINGFORRELEASE'";
                rs = sql.execute(strQuery)
                if (!rs.EOF) {

            %>
            <span style="font-weight: bold;"><% = rs(0) %></span>
            <%
     
                     //Response.Write(rs(0)+ "&nbsp;&nbsp;");
                     Response.Write("&nbsp;&nbsp;&nbsp;");

                     rs.movenext();
                 }
            %>

            <!--  <div class="letter_to_left" style="float: left;"> WaitingForProduction:  </div>-->
            WaitingForProduction:

            <img src="allinfos_status_open.png" />

            <%
                strQuery = "select count(*)  from SlitteApp where (CertificateState = 0  OR  CertificateState = Null) AND Status ='WAITINGFORPRODUCTION'";
                rs = sql.execute(strQuery)
                if (!rs.EOF) {
            %>
            <span style="font-weight: bold;"><% = rs(0) %></span>
            <%
             
                    //Response.Write(rs(0)+ "&nbsp;&nbsp;");
                    Response.Write("&nbsp;");

                    rs.movenext();
                }
            %>

            <img src="allinfos_status_pending.png" />
            <% 
                strQuery = "select count(*)  from SlitteApp where CertificateState = 1  AND  Status ='WAITINGFORPRODUCTION'";
                rs = sql.execute(strQuery)
                if (!rs.EOF) {

            %>
            <span style="font-weight: bold;"><% = rs(0) %></span>
            <%
             
                    //Response.Write(rs(0));
                    Response.Write( "&nbsp;");
                    rs.movenext();
                }
            %>
            <img src="allinfos_status_ready.png" />
            <%
                strQuery = "select count(*)  from SlitteApp where CertificateState = 2 AND  Status ='WAITINGFORPRODUCTION'";
                rs = sql.execute(strQuery)
                if (!rs.EOF) {

            %>
            <span style="font-weight: bold;"><% = rs(0) %></span>
            <%
                    //Response.Write(rs(0) + "&nbsp; &nbsp;");
                    Response.Write("&nbsp; &nbsp;&nbsp;");
                    rs.movenext();
                }
            %>


            <!--<div class="letter_to_left" style="float: left;"> Rejected: </div>-->
            Rejected:
             <%
                 strQuery = "Select count(*)  from SlitteApp where Status ='INPRODUCTION'";
                 rs = sql.execute(strQuery)
                 if (!rs.EOF) {

             %>
            <span style="font-weight: bold;"><% = rs(0) %></span>
            <%
                     //Response.Write(rs(0) + " ");
                     rs.movenext();
                 }
            %>
           <br/>

            <div style="font-size: 8pt;"> (*Alle Datumswerten größer als 2 Tage, werden rot dargestellt. **Am Montags [Apps Gestern] sind die Summe: Freitag, Samstag und Sonntag.)</div>
             
        </div>



    </div>



</body>
</html>
