import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class samplepcfcontrol implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    /**
     * Empty constructor.
     */
     private _value: number;
     private _notifyOutputChanged: () => void;
     private _container: HTMLDivElement;
     private _titleContainer: HTMLDivElement;
     private _title : HTMLDivElement;
      
  
     private _range : HTMLInputElement;
     private _sheet : HTMLStyleElement;
     private _spanValue : HTMLSpanElement;
     private _refreshData: EventListenerOrEventListenerObject;
     private _context : ComponentFramework.Context<IInputs>;
     /**
      * Empty constructor.
      */
     constructor()
     {
  
     }
  
     /**
      * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
      * Data-set values are not initialized here, use updateView.
      * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
      * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
      * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
      * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
      */
     public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
     {
         this._context=context;
         this._notifyOutputChanged = notifyOutputChanged;
         this._refreshData = this.refreshData.bind(this);
  
         this._container= document.createElement("div");
         this._container.setAttribute("class","slidecontainer");
         this._titleContainer= document.createElement("div");
         this._titleContainer.setAttribute("class","title-container");
  
         // CRM attributes bound to the control properties. 
          // @ts-ignore 
         var optionsetAttribute: string = context.parameters.annualSales.attributes.LogicalName;
         //@ts-ignore
         var options = window.parent.Xrm.Page.getAttribute(optionsetAttribute).getOptions();
         var width: number = 100/options.length;
         //@ts-ignore
         for(var i:number = 0; i < options.length; i++)
         {
             this._title =document.createElement("div");
             this._title.setAttribute("class","tile");
             this._title.innerText =options[i].text?options[i].text:"Blank";
             this._title.style.width=(width + "%").toString();
             this._titleContainer.appendChild(this._title);  
                      
         }
         this._container.appendChild(this._titleContainer);
  
         this._range=document.createElement("input");
         this._range.setAttribute("type","range");
         this._range.addEventListener("input", this._refreshData);
         this._range.setAttribute("min","100000000");
         this._range.setAttribute("max","100000005")
         this._range.setAttribute("class","slider")
  
         this._container.appendChild(this._range);
  
         this._sheet = document.createElement('style')
         this._sheet.innerHTML = ".slider::-webkit-slider-thumb { width: 20%;}";
         document.body.appendChild(this._sheet);
  
         this._spanValue =document.createElement("span");
         this._container.appendChild(this._spanValue);
         container.appendChild(this._container);
         this._value=context.parameters.annualSales.raw?context.parameters.annualSales.raw:0;
         this._range.value = (context.parameters.annualSales.raw?context.parameters.annualSales.raw:0).toString();
         //this._spanValue.innerHTML = (context.parameters.annualSales.formatted?context.parameters.annualSales.formatted:"0").toString();
  
          
        
  
          //@ts-ignore 
         var crmTagStringsAttributeValue = window.parent.Xrm.Page.getAttribute(optionsetAttribute).getText();
           //@ts-ignore
         this._spanValue.innerHTML = crmTagStringsAttributeValue;
          
  
     }
     public refreshData(evt: Event): void {
          
         this._spanValue.innerHTML = this._range.value;
         this._notifyOutputChanged();
       }
  
  
     /**
      * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
      * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
      */
     public updateView(context: ComponentFramework.Context<IInputs>): void
     {
         this._context = context;
         this._value = context.parameters.annualSales.raw?context.parameters.annualSales.raw:0;
         this._range.value = (context.parameters.annualSales.raw?context.parameters.annualSales.raw:0).toString();
         //this._spanValue.innerHTML = (context.parameters.annualSales.formatted?context.parameters.annualSales.formatted:0).toString();
         // CRM attributes bound to the control properties. 
         // @ts-ignore 
         var crmTagStringsAttribute = context.parameters.annualSales.attributes.LogicalName;
  
         // @ts-ignore 
         var crmTagStringsAttributeValue = window.parent.Xrm.Page.getAttribute(crmTagStringsAttribute).getText();
          // @ts-ignore
         this._spanValue.innerHTML = crmTagStringsAttributeValue;
     }
  
     /** 
      * It is called by the framework prior to a control receiving new data. 
      * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
      */
     public getOutputs(): IOutputs
     {
         return {annualSales:parseInt(this._range.value)};
     }
  
     /** 
      * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
      * i.e. cancelling any pending remote calls, removing listeners, etc.
      */
     public destroy(): void
     {
         this._range.removeEventListener("input", this._refreshData);
     }
 }