<%@ Page Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="Home.aspx.cs" Inherits="home"%>

<%@ Register Assembly="RBGP.WebControls.EmailLink" Namespace="RBGP.WebControls" TagPrefix="rbgp" %>
<%@ Register Assembly="Thumb2Large" Namespace="RBGP.WebControls" TagPrefix="rbgp"%>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <a name="s0"></a>
    <!--
    <a href="http://www.codeup.co.uk/large_picture.html?title=Orchard retreat from above&image=http://orchardretreat.norrisandson.co.uk/images/houseFromAbove_w600.jpg&url=http://orchardretreat.norrisandson.co.uk/home.aspx#s0">
        <img alt="Orchard retreat from above" src="images/thumbs/houseFromAbove_w600.jpg" style="float:right; margin:15px" />
    </a>
    -->
    
        
    <div class="centre">
        <p>Relax and enjoy the tranquillity of this superb retreat.</p>
        <img title="Orchard retreat from above" alt="Orchard retreat from above" src="images/houseFromAbove_w480_new.jpg"/>
    </div>
    <div id="contentText">
    <p>Welcome to our <b>Orchard Retreat</b>, situated in Marden, within the low Kentish Weald, just 8 miles south of Maidstone and with Royal Tunbridge Wells within easy driving distance to the west.  There  is a regular train service to London Charing Cross from the station in the village, 1 mile away.</p>
    <p>We are a non-smoking establishment and unfortunately cannot accommodate children or pets.</p>
    <p>This completely self-contained, picturesque studio, situated within beautiful landscaped gardens and adjacent to pear orchards is  close to <b>Sissinghurst Castle Garden, Leeds Castle, Bodiam Castle</b> and many other National Trust properties.  For garden lovers there are numerous National Garden Scheme locations nearby and of course for walkers or cyclists, there is the Bedgebury Pinetum.</p>
    <a name="s1"></a>
    <!-- 
    <a href="http://www.codeup.co.uk/large_picture.html?title=Front of studio&image=http://orchardretreat.norrisandson.co.uk/images/houseFront_w600.jpg&url=http://orchardretreat.norrisandson.co.uk/home.aspx#s0">
        <img src="images/thumbs/houseFront_w600.jpg" alt="front of studio" style="float:left; margin-right:15px" />
    </a>
    -->
    <div class="centre">
    <rbgp:Thumb2LargeControl ID="frontOfStudio" runat="server" Title="Front of studio" Section="0" ImageUrl="images/housefront_w600_new.jpg"></rbgp:Thumb2LargeControl>
    </div>
    <p class="centre">TARIFF 2009</p>
<!-- <a href="http:\\www.leeds-castle.com" target="_blank"> -->
    <div id="tariff">
        <ul>
            <li>Bed with Breakfast Pack in room..  ..  ..  £65 per night<br />( Single occupancy £55 per night )</li>
            <!-- <li>10% Discount offered for multiple night stays</li> -->
            <li>Full English Breakfast extra at £4.50 per person</li>
            <li>A non-refundable booking deposit of £30 is required unless cancellation is received 48 hours prior to arrival. Payment may be made by Cash or Cheque (in sterling) with a Banker's Card.  Unfortunately we do not have credit card facilities.</li>
            <li>We supply a breakfast pack of cereals, bread, local preserves and yoghurt to your room or a <b>Full English Breakfast</b> can be taken by arrangement in the dining room of the main house, a 400 year old Kentish Farmhouse.</li>        
        </ul>
     </div>

    <!--
    <a href="http://www.codeup.co.uk/large_picture.html?title=Bedroom&image=http://orchardretreat.norrisandson.co.uk/images/bedroom_w600.jpg&url=http://orchardretreat.norrisandson.co.uk/home.aspx#s1">
        <img src="images/thumbs/bedroom_w600.jpg" alt="bed" style="float:right; padding-top:10px;" /> 
    </a>
    -->
    <div class="centre">
        <img src="images/bedroom_w600.jpg" alt="Bedroom" id="bedroom" />
    </div>
    <!--
    <rbgp:Thumb2LargeControl runat="server" ID="bedroom" Title="Bedroom" Section="1" ImageUrl="images/bedroom_w600.jpg" style="float:right; padding-top:10px;"></rbgp:Thumb2LargeControl>
    -->
    <a name="s2"></a> 
    <ul class="facilities">
        <li>Double Room en-suite with facilities</li>
        <li>Bathrobes supplied</li>
        <li>Full air conditioning system</li>
        <li>TV with DVD player</li>
        <li>Tea and coffee making facilities</li>
        <li>Breakfast pack supplied to your room</li>
        <li>Fridge with mineral water/ice making facilities</li>
    </ul>
    <br /><br />
    
    <p class="centre">
        <rbgp:Thumb2LargeControl runat="server" ID="shower1" Title="Shower (picture 1)" Section="2" ImageUrl="images/shower1.jpg"></rbgp:Thumb2LargeControl>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <rbgp:Thumb2LargeControl runat="server" ID="shower2" Title="Shower (picture 2)" Section="2" ImageUrl="images/shower2.jpg"></rbgp:Thumb2LargeControl>
    </p>
    
        <div class="centre">
            <!-- 
            <a href="http://www.codeup.co.uk/large_picture.html?title=Shower (picture 2)&image=http://orchardretreat.norrisandson.co.uk/images/shower2.jpg&url=http://orchardretreat.norrisandson.co.uk/home.aspx#s2">
                <img src="images/thumbs/shower2.jpg" alt="shower photo 2" style="float:right; margin:20px" />
            </a>
            -->
            
            <p>Your hostess Lynda Treliving looks forward to meeting you.</p>    
            <p>Please telephone: 01622 831908</p>
            <p>OR</p>
            <p>Mob: 07814 174661</p>
            <p>e-mail: <rbgp:EmailLink ID="emailId" runat="server" NavigateUrl="mailto:BandB@orchardretreat.co.uk">BandB@orchardretreat.co.uk</rbgp:EmailLink></p>
        </div>

    </div>
</asp:Content>

