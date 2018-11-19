<?xml version="1.0" encoding="utf-8"?>
<applicationModel>
  <name>Application Model</name>
  <className>CastorObjects</className>
  <description>All Castor objects will store here</description>
  <objects>
    <object technology="Standard Windows" type="Dialog">
      <name>Message from webpage</name>
      <codeName>MessageFromWebpageDialog</codeName>
      <identification>
        <properties>
          <property name="IsOwnedWindow" type="boolean">true</property>
          <property name="IsChildWindow" type="boolean">false</property>
          <property name="Text" type="string">Message from webpage</property>
          <property name="NativeClass" type="string">#32770</property>
        </properties>
      </identification>
      <children>
        <object technology="Standard Windows" type="Button">
          <name>OK</name>
          <codeName>OKButton</codeName>
          <identification>
            <properties>
              <property name="Text" type="string">OK</property>
              <property name="NativeClass" type="string">Button</property>
            </properties>
          </identification>
        </object>
      </children>
    </object>
    <object technology="Web" type="Page">
      <name>Castor</name>
      <codeName>Castor</codeName>
      <identification>
        <properties>
          <property name="Title" type="string">Castor</property>
        </properties>
      </identification>
      <children>
        <object technology="Web" type="Button">
          <name>SEBO PD Buyer</name>
          <codeName>SEBOPDBuyerButton</codeName>
          <identification>
            <properties>
              <property name="ButtonType" type="string">button</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">SEBO PD Buyer</property>
            </properties>
          </identification>
        </object>
        <object technology="Web" type="Button">
          <name>CNHK PD Merch</name>
          <codeName>CNHKPDMerchButton</codeName>
          <identification>
            <properties>
              <property name="ButtonType" type="string">button</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">CNHK PD Merch</property>
            </properties>
          </identification>
        </object>
      </children>
    </object>
    <object technology="Web" type="Page">
      <name>CastorPage2</name>
      <codeName>CastorPage2</codeName>
      <identification>
        <properties>
          <property name="Title" type="string">Castor</property>
        </properties>
        <index>0</index>
      </identification>
      <children>
        <object technology="Web" type="Frame">
          <name>content</name>
          <codeName>ContentFrame</codeName>
          <identification>
            <properties>
              <property name="Name" type="string">content</property>
            </properties>
          </identification>
          <children>
            <object technology="Web" type="Link">
              <name>Product Developments</name>
              <codeName>ProductDevelopmentsLink</codeName>
              <identification>
                <properties>
                  <property name="TagName" type="string">A</property>
                  <property name="InnerText" type="string">Product Developments</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="Web" type="Frame">
          <name>tvcTabs0_HMProductDevelopmentBuyerStartPageTabcontentFrame</name>
          <codeName>TvcTabs0HMProductDevelopmentBuyerStartPageTabcontentFrame</codeName>
          <identification>
            <properties>
              <property name="Name" type="string">tvcTabs0_HMProductDevelopmentBuyerStartPageTabcontentFrame</property>
            </properties>
          </identification>
          <children>
            <object technology="Web" type="Link">
              <name>Product Developments</name>
              <codeName>ProductDevelopmentsLink</codeName>
              <identification>
                <properties>
                  <property name="TagName" type="string">A</property>
                  <property name="InnerText" type="string">Product Developments</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="Web" type="Frame">
          <name>tableContentFrame</name>
          <codeName>TableContentFrame</codeName>
          <identification>
            <properties>
              <property name="Name" type="string">tableContentFrame</property>
            </properties>
          </identification>
          <children>
            <object technology="Web" type="WebElement">
              <name>Create Product Development</name>
              <codeName>CreateProductDevelopmentWebElement</codeName>
              <identification>
                <properties>
                  <property name="TagName" type="string">SPAN</property>
                  <property name="InnerText" type="string">Create Product Development</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
      </children>
    </object>
    <object technology="Web" type="Page">
      <name>Page</name>
      <codeName>Page</codeName>
      <identification>
        <properties>
          <property name="Title" type="string" />
        </properties>
      </identification>
      <children>
        <object technology="Web" type="WebElement">
          <name>Oops I closed Castor, open it again please.</name>
          <codeName>OopsIClosedCastorOpenItAgainPleaseWebElement</codeName>
          <identification>
            <properties>
              <property name="ClassName" type="string" />
              <property name="TagName" type="string">TD</property>
              <property name="InnerText" type="string" />
            </properties>
          </identification>
        </object>
      </children>
    </object>
  </objects>
</applicationModel>