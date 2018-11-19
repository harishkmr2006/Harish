<?xml version="1.0" encoding="utf-8"?>
<applicationModel>
  <name>HMORDER</name>
  <className>HMOrderObject</className>
  <objects>
    <object technology="WinForms" type="Window">
      <name>HMOrder</name>
      <codeName>HMOrderWindow</codeName>
      <identification>
        <properties>
          <property name="IsChildWindow" type="boolean">false</property>
          <property name="IsOwnedWindow" type="boolean">false</property>
          <property name="ObjectName" type="string">OrderParent</property>
          <property name="WindowTitleRegExp" type="string">HMOrder</property>
        </properties>
      </identification>
      <children>
        <object technology="WinForms" type="UiObject">
          <name>_BaseForm_Toolbars_Dock_Area_Top</name>
          <codeName>BaseFormToolbarsDockAreaTopUiObject</codeName>
          <identification>
            <properties>
              <property name="ObjectName" type="string">_BaseForm_Toolbars_Dock_Area_Top</property>
              <property name="FullNamePath" type="string">_BaseForm_Toolbars_Dock_Area_Top;OrderParent</property>
              <property name="FullType" type="string">Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea</property>
            </properties>
          </identification>
        </object>
        <object technology="WinForms" type="Window">
          <name>Order Explorer</name>
          <codeName>OrderExplorerWindow</codeName>
          <identification>
            <properties>
              <property name="IsChildWindow" type="boolean">true</property>
              <property name="IsOwnedWindow" type="boolean">false</property>
              <property name="ObjectName" type="string">OrderExplorer</property>
              <property name="WindowTitleRegExp" type="string">Order Explorer</property>
            </properties>
          </identification>
          <children>
            <object technology="WinForms" type="UiObject">
              <name>ultraTreeDepartments</name>
              <codeName>UltraTreeDepartmentsUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraTreeDepartments</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="WinForms" type="Window">
          <name>New Order</name>
          <codeName>NewOrder</codeName>
          <identification>
            <properties>
              <property name="IsChildWindow" type="boolean">false</property>
              <property name="IsOwnedWindow" type="boolean">true</property>
              <property name="ObjectName" type="string">CreateOrderWizardForm</property>
            </properties>
          </identification>
          <children>
            <object technology="WinForms" type="Table" subtype="UltraWinGrid">
              <name>gridBoardCards</name>
              <codeName>GridBoardCardsTable</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">gridBoardCards</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="UiObject">
              <name>cboCurveGroups</name>
              <codeName>CboCurveGroupsUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">cboCurveGroups</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="UiObject">
              <name>chkShowAll</name>
              <codeName>ChkShowAllUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">chkShowAll</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Button">
              <name>Next</name>
              <codeName>NextButton</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">NextButton</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="UiObject">
              <name>cboSizeScales</name>
              <codeName>CboSizeScalesUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">cboSizeScales</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Table" subtype="UltraWinGrid">
              <name>optionsGrid</name>
              <codeName>OptionsGridTable</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">optionsGrid</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Button">
              <name>Finish</name>
              <codeName>FinishButton</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">FinishButton</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="WinForms" type="Window">
          <name>OrderEditer</name>
          <codeName>OrderEditProperties</codeName>
          <identification>
            <properties>
              <property name="IsChildWindow" type="boolean">true</property>
              <property name="IsOwnedWindow" type="boolean">false</property>
              <property name="ObjectName" type="string">OrderEditor</property>
            </properties>
          </identification>
          <children>
            <object technology="WinForms" type="UiObject">
              <name>Order Properties</name>
              <codeName>OrderPropertiesUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraGroupBox1</property>
                  <property name="FullNamePath" type="string">ultraGroupBox1;orderPropertiesControl1;PropertiesGroupBox;dockableWindow3;windowDockingArea1;OrderEditor;;OrderParent</property>
                  <property name="FullType" type="string">Infragistics.Win.Misc.UltraGroupBox</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Table" subtype="UltraWinGrid">
              <name>ultraGridMainOrder</name>
              <codeName>UltraGridMainOrderTable</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraGridMainOrder</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Table" subtype="UltraWinGrid">
              <name>ultraGridValidations</name>
              <codeName>UltraGridValidationsTable</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraGridValidations</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="UiObject">
              <name>Preliminary</name>
              <codeName>PreliminaryUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">cboStatus</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="UiObject">
              <name>ultraTabControl1</name>
              <codeName>UltraTabControl1UiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraTabControl1</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Table" subtype="UltraWinGrid">
              <name>ultraGridPriceCalculation</name>
              <codeName>UltraGridPriceCalculationTable</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraGridPriceCalculation</property>
                </properties>
              </identification>
            </object>
            <object technology="WinForms" type="Table" subtype="UltraWinGrid">
              <name>ultraGridTMQMAQ</name>
              <codeName>UltraGridTMQMAQTable</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">ultraGridTMQMAQ</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="WinForms" type="Window">
          <name>Update TOD</name>
          <codeName>UpdateTODWindow</codeName>
          <identification>
            <properties>
              <property name="IsChildWindow" type="boolean">false</property>
              <property name="IsOwnedWindow" type="boolean">true</property>
              <property name="ObjectName" type="string">ConfirmationDialog</property>
              <property name="WindowTitleRegExp" type="string">Update TOD</property>
            </properties>
          </identification>
          <children>
            <object technology="WinForms" type="UiObject">
              <name>Yes</name>
              <codeName>YesUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">btnYes</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="WinForms" type="ToolBar">
          <name>_ultraGridRightClickMenu</name>
          <codeName>UltraGridRightClickMenuToolBar</codeName>
          <identification>
            <properties>
              <property name="ObjectName" type="string">_ultraGridRightClickMenu</property>
            </properties>
          </identification>
        </object>
        <object technology="WinForms" type="Window">
          <name>ProductWindow</name>
          <codeName>TestAutoProdWindow</codeName>
          <identification>
            <properties>
              <property name="IsChildWindow" type="boolean">true</property>
              <property name="IsOwnedWindow" type="boolean">false</property>
              <property name="ObjectName" type="string">OrderEditor</property>
            </properties>
          </identification>
          <children>
            <object technology="WinForms" type="EditField">
              <name>txtOrderNo_EmbeddableTextBox</name>
              <codeName>TxtOrderNoEmbeddableTextBoxEditField</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">txtOrderNo_EmbeddableTextBox</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
        <object technology="WinForms" type="Window">
          <name>Closing HMOrder</name>
          <codeName>ClosingHMOrderWindow</codeName>
          <identification>
            <properties>
              <property name="IsChildWindow" type="boolean">false</property>
              <property name="IsOwnedWindow" type="boolean">true</property>
              <property name="ObjectName" type="string">ConfirmationDialog</property>
              <property name="WindowTitleRegExp" type="string">Closing HMOrder</property>
            </properties>
          </identification>
          <children>
            <object technology="WinForms" type="UiObject">
              <name>Yes</name>
              <codeName>YesUiObject</codeName>
              <identification>
                <properties>
                  <property name="ObjectName" type="string">btnYes</property>
                </properties>
              </identification>
            </object>
          </children>
        </object>
      </children>
    </object>
    <object technology="Standard Windows" type="Window">
      <name>Mozilla Firefox</name>
      <codeName>MozillaFirefoxWindow</codeName>
      <identification>
        <properties>
          <property name="IsOwnedWindow" type="boolean">false</property>
          <property name="IsChildWindow" type="boolean">false</property>
          <property name="Text" type="string">Mozilla Firefox</property>
          <property name="WindowClassRegExp" type="string">MozillaWindowClass</property>
          <property name="WindowTitleRegExp" type="string">Mozilla Firefox</property>
        </properties>
      </identification>
      <children>
        <object technology="Standard Windows" type="Dialog">
          <name>Authentication Required</name>
          <codeName>AuthenticationRequiredDialog</codeName>
          <identification>
            <properties>
              <property name="IsOwnedWindow" type="boolean">true</property>
              <property name="IsChildWindow" type="boolean">false</property>
              <property name="Text" type="string">Authentication Required</property>
              <property name="NativeClass" type="string">MozillaDialogClass</property>
            </properties>
          </identification>
        </object>
      </children>
    </object>
  </objects>
</applicationModel>