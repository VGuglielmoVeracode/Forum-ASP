<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2019 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  https://www.webwiz.net/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 18, The Glenmore Centre, Fancy Road, Poole, Dorset, BH12 4FB, England
'**  https://www.webwiz.net
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************






'Set the response buffer to true as we maybe redirecting
Response.Buffer	= True


'Clear server objects
Call closeDatabase()


'If RSS is not enabled send the user away
If blnRSS = False Then

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'Set the content type for feed
Response.ContentType = "application/xml"

%><?xml version="1.0" encoding="<% = strPageEncoding %>" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN" />
<xsl:variable name="title" select="/rss/channel/title" />
<xsl:template match="/">
<xsl:element name="html">
<head>
 <title><xsl:value-of select="$title"/></title>
 <script type="text/javascript" src="includes/rss_disableOutputEscaping.js" />
 <link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" media="all"/>
 <link href="<% = strCSSfile %>rss_style.css" rel="stylesheet" type="text/css" media="all"/>
</head>
<xsl:apply-templates select="rss/channel"/>
</xsl:element>
</xsl:template>
<xsl:template match="channel">
<body onload="go_decoding();">
 <div id="cometestme" style="display:none;">
  <xsl:text disable-output-escaping="yes">&amp;amp;</xsl:text>
 </div>
 <table class="rssContainer">
  <tr>
   <td>
    <xsl:apply-templates select="title"/>
    <xsl:apply-templates select="description"/>
    <div class="contentBlock">
     <xsl:apply-templates select="item"/>
    </div>
    <br />
    <span class="rssCopy">RSS Online Forum Viewer v1.2<br /><xsl:value-of select="copyright"/></span>
   </td>
  </tr>
 </table>
</body>
</xsl:template>


<xsl:template match="title">
 <div class="bannerBlock">
  <span style="float:right;"><xsl:apply-templates select="../image"/></span>
  <span class="rssTitle"><a href="{link}"><xsl:value-of select="$title"/></a></span>
  <br /><% = strTxtSyndicatedForumContent %>
  <br /><br />
 </div>
 <br />
</xsl:template>


<xsl:template match="description">
 <xsl:variable name="feedUrl" select="/rss/channel/WebWizForums:feedURL" xmlns:WebWizForums="https://syndication.webwiz.net/rss_namespace/"/>
 <div class="headerBlock">
  <div class="subscribeBlock">
   <strong><% = strTxtSubscribeNow %></strong>
   <br /><br />...<% = strTxtSubscribeWithWebBasedRSS %>:-
   <a href="http://fusion.google.com/add?feedurl=http://{$feedUrl}" class="rssButton"><img src="<% = strImagePath %>rss_google.gif" alt="Add to Google"/></a>
   <a href="http://www.newsgator.com/ngs/subscriber/subext.aspx?url=http://{$feedUrl}" class="rssButton"><img src="<% = strImagePath %>rss_newsgator.gif" alt="Subscribe in NewsGator"/></a>
   <a href="http://add.my.yahoo.com/rss?url=http://{$feedUrl}" class="rssButton"><img src="<% = strImagePath %>rss_yahoo.gif" alt="Add to my Yahoo"/></a>
   <xsl:element name="a">
    <xsl:attribute name="href">http://client.pluck.com/pluckit/prompt.aspx?GCID=C12286x053&amp;a=http://<xsl:value-of select="$feedUrl"/>&amp;t=<xsl:value-of select="$title"/></xsl:attribute>
    <xsl:attribute name="class">rssButton</xsl:attribute>
    <img src="<% = strImagePath %>rss_pluck.png" alt="Subscribe with Pluck RSS reader" border="0" />
   </xsl:element>
   <a href="http://www.rojo.com/add-subscription?resource=http://{$feedUrl}" class="rssButton"><img src="<% = strImagePath %>rss_rojo.gif" alt="Subscribe in Rojo"/></a>
   <br /><br />...<% = strTxtWithOtherReaders %>:-
   <br />
   <a href="feed://{$feedUrl}"><xsl:value-of select="$title"/></a>
  </div>
  <br /><strong><xsl:value-of select="."/>.</strong>
  <br />
  <br /><% = strTxtThisRSSFileIntendedToBeSyndicated %>
  <br /><br />
  <a href="<% = strForumPath %>help.asp#FAQ29" target="_blank"><% = strTxtWhatIsAnRSSFeed %></a>
  <br /><br /><br /><br />
  <span class="rssHeading"><% = strTxtCurrentFeedContent %></span>
 </div>
</xsl:template>


 <xsl:template match="item">
  <ul xmlns="http://www.w3.org/1999/xhtml">
   <li class="RSSbullet">
    <a href="{link}" class="rssLink" target="_blank"><xsl:value-of select="title" /></a>
    <br />
    <span class="rssPostedDate"><xsl:if test="count(child::pubDate)=1"><xsl:value-of select="substring(pubDate,5)"/></xsl:if></span>
    <br /><br />
    <div class="itemcontent" name="decodeable" style="overflow:auto;">
     <xsl:value-of select="description" disable-output-escaping="yes" />
    </div>
    <br />
    <hr />
    <br />
   </li>
  </ul>
 </xsl:template>


 <xsl:template match="image">
  <a href="{link}">
   <xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
    <xsl:attribute name="src">
     <xsl:value-of select="url"/>
    </xsl:attribute>
    <xsl:attribute name="alt">Link to <xsl:value-of select="title"/></xsl:attribute>
    <xsl:attribute name="id">feedimage</xsl:attribute>
   </xsl:element>
  </a>
  <xsl:text/>
 </xsl:template>


</xsl:stylesheet>