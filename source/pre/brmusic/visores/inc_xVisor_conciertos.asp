<!--#include file="inc_conn.asp" -->

<style type="text/css">

/* Enlaces
----------------------------------------------- */
	a:link {
		color:#9db;
		}
	a:visited {
		color:#798;
		}
	a:hover {
		color:#fff;
		}
	a img {
		border-width:0;
		}

/* General
----------------------------------------------- */

	#content {
		width:538px;
		margin:0 auto;
		text-align:left;
		}
	.main {
		width:375px;
		float:left;
		margin:15px 0 0;
		padding:0;
		line-height:1.5em;
		background: #F3F3F3 url(/arch/esquinas_principal_botton.gif) no-repeat left bottom;
		}
	.main2 {
		float:left;
		width:100%;
		padding:10px 0 0;
		background: url(/arch/esquinas_principal_top.gif) no-repeat left top;
		}
		
/* Barra lateral
----------------------------------------------- */

	#sidebar {
		width:150px;
		float:right;
		margin:15px 0 0;
		font-size:97%;
		line-height:1.5em;
		}
	.box {
		margin:0 0 15px;
		padding:10px 0 0;
		color:#abc;
		background-color: #000000;
		background-image: url(/arch/b-lateral_top.gif);
		background-repeat: no-repeat;
		background-position: left top;
		}
	.box2 {
		background:url("/arch/b-lateral_bot.gif") no-repeat left bottom;
		padding:0 13px 8px;
		}

	.sidebar-title {
		margin:0;
		padding:0 0 .2em;
		border-bottom:1px dotted #456;
		font-size:115%;
		line-height:1.5em;
		color:#abc;
		}
	.box ul {
		margin:.5em 0 1.25em;
		padding:0 0px;
		list-style:none;
		}
	.box ul li {
		background:url("http://www.blogblog.com/rounders3/icon_arrow_sm.gif") no-repeat 2px .25em;
		margin:0;
		padding:0 0 3px 16px;
		margin-bottom:3px;
		border-bottom:1px dotted #345;
		line-height:1.4em;
		}
	.box p {
		margin:0 0 .6em;
		}

	.concierto {
		margin:.3em 0 25px;
		padding:0 13px;
		}
	.concierto-titulo {
		display:block;
		margin:0;
		font-size:135%;
		line-height:1.5em;
		background:url("http://www.blogblog.com/rounders3/icon_arrow.gif") no-repeat 10px .5em;
		padding:2px 14px 2px 29px;
		color:#CC3300;
		border-bottom-width: thin;
		border-bottom-style: dotted;
		border-bottom-color: #666666;
		}
	.concierto-cuerpo {
		display:block;
		margin:0;
		line-height:1.5em;
		padding:6px 10px 2px 15px;
		}
	#main .concierto-titulo strong {
		line-height:1.5em;
		display:block;
		color:#316AC5;
		padding-top: 2px;
		padding-right: 14px;
		padding-bottom: 2px;
		padding-left: 29px;
		background-image: url(http://www.estadologico.com/nenito/arch/flecha-link.gif);
		background-repeat: no-repeat;
		background-position: 13px 0.5em;
		margin-bottom: 10px;
		}
	.concierto-body {
		padding:10px 14px 10px 29px;
		}
	.concierto img {
		margin:0 0 5px 0;
		padding:4px;
		border:1px solid #586;
		}


</style>
</head><body>
<div id="content">
<div id="sidebar">
    <div class="box">
      <div class="box2">
        <div class="box3">
          <h2 class="sidebar-title">Enlaces</h2>
          <ul>
            <li><a href="http://www.podcast-es.org/" target="_blank">Podcast-es</a><br />
              Proyecto colaborativo en el que se da cabida a todos las personas hispano-hablantes interesadas en el mundo del Podcasting.<br />
            </li>
          </ul>
        </div>
      </div>
    </div>
  </div>

<%
	sql = "SELECT * FROM REGISTROS"
	set re = Server.CreateObject("ADODB.Recordset") : re.ActiveConnection = conn_ : re.Source = sql : re.CursorType = 3 : re.CursorLocation = 2 : re.LockType = 3
	re.Open()
	
	if not re.eof then
		while not re.eof%>
			<div class="main">
				<div class="main2">
					<div class="concierto">
						<a href="#" title=""> <h3 class="concierto-titulo"><%=re("R_TITULO")%></h3> </a>
						<div class="concierto-cuerpo">
							<div id="fecha"> <%=re("R_TEXT1")%> </div>
							<div id="lugar"> <%=re("R_TEXT2")%> </div>
						</div>
					</div>
				</div>
			</div>
			<%re.movenext
		wend
	end if

	re.Close()
	set re = Nothing
%>
</div>
