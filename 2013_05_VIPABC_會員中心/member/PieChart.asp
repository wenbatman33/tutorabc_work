<!DOCTYPE html>
<html>
    <head>
        <title></title>
        <link class="include" rel="stylesheet" type="text/css" href="/css/chart/jquery.jqplot.min.css" />
        <script class="include" type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery-1.6.2.min.js"></script>
      
        <!--[if lt IE 19]><script language="javascript" type="text/javascript" src="/lib/javascript/chart/excanvas.js"></script><![endif]-->
        <script class="include" type="text/javascript" src="/lib/javascript/chart/jquery.jqplot.min.js"></script>
        <script class="include" type="text/javascript" src="/lib/javascript/chart/jqplot.pieRenderer.min.js"></script>
    </head>
    <body>
        <div id="charts"></div>

        <script type="text/javascript">
        
         
        $(document).ready(function(){
            var data = [
                ['<%=request("w1")%>', <%=request("n1")%>],
                ['<%=request("w2")%>', <%=request("n2")%>], 
                ['<%=request("w3")%>', <%=request("n3")%>],     
                ['<%=request("w4")%>', <%=request("n4")%>]
          ];
          var plot1 = jQuery.jqplot ('charts', [data], 
            { 
              seriesDefaults: {
                // Make this a pie chart.
                renderer: jQuery.jqplot.PieRenderer, 
                rendererOptions: {
                  // Put data labels on the pie slices.
                  // By default, labels show the percentage of the slice.
                  showDataLabels: true
                }
              }, 
              legend: { show:true, location: 'e' }
            }
          );
        });
        </script>
    </body>
</html>