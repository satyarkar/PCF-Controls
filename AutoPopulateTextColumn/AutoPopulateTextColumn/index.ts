import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class AutoPopulateTextColumn implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	// reference to Power Apps component framework Context object
	private _context: ComponentFramework.Context<IInputs>;
	// reference to the component container HTMLDivElement
	private _container: HTMLDivElement;
	// power Apps component framework delegate which will be assigned to this object which would be called whenever any update happens.
	private _notifyOutputChanged: () => void;
	// input element that is used to auto populate
	private inputElement: HTMLInputElement;
	// value of the column is stored and used inside the control
	private _value: string;
	private _regardingValue: any;
	private _configValue: string;
	/**
	 * Empty constructor.
	 */
	constructor() { }

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement) {
		this._context = context;
		this._container = document.createElement("div");
		this._notifyOutputChanged = notifyOutputChanged;
		// creating HTML elements for the auto populate text column
		this.inputElement = document.createElement("input");
		this.inputElement.setAttribute("class", "boundText");
		this.inputElement.setAttribute("placeholder", "---");
		// retrieving the latest value from the component.
		this._value = context.parameters.boundTextProperty.raw
		? context.parameters.boundTextProperty.raw
		: "";
		this.inputElement.value = this._value;
		this.inputElement.addEventListener("blur", this.onBlur.bind(this));
		this._regardingValue = context.parameters.regardingValue.raw
		? context.parameters.regardingValue.raw
		: null;
		this._configValue = context.parameters.configValue.raw
		? context.parameters.configValue.raw
		: "";
		this.autoPopulate(this._regardingValue);

		// appending the HTML elements to the component's HTML container element.
		this._container.appendChild(this.inputElement);
		container.appendChild(this._container);
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		let newRegardingValue: any = context.parameters.regardingValue.raw
		? context.parameters.regardingValue.raw
		: null;
		this.autoPopulate(newRegardingValue);
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			boundTextProperty: this._value
		};
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// add code to cleanup control if necessary
	}

	/**
	 * Called when the text column value is changed
	 */
	private onBlur(): void {
		this._value = this.inputElement.value;
		this._notifyOutputChanged();
	}

	/**
	 * Called this function to auto populate the formatted test.
	 * @param newValue The new value of type Lookup.Regarding.
	 */
	autoPopulate(newValue: any): void {
		this._regardingValue = newValue;
		if(newValue === null || newValue.length === 0) {
			this.notify("");
		} else if(this.inputElement.value === "") {
			this.getFormattedText(newValue[0].entityType, newValue[0].id );
		}
	}

	/**
	 * Called this function to notify the framework when text value is changed.
	 * @param newText The new formmated text.
	 */
	private notify(newText:any):void {
		this.inputElement.value = newText;
		this._value = this.inputElement.value;
		this._notifyOutputChanged();
	}

	/**
	 * Called this function to retrive the regarding record and format the output.
	 * @param tableType The table type of selected Lookup.Regarding.
	 * @param id The table id of selected Lookup.Regarding.
	 */
	getFormattedText(tableType: string, id: string): void {
		let output:string = "";
		let selectedTypeConfig: string[] = [];
		let multiTypeConfig: string[] = this._configValue.split("^");
		if(multiTypeConfig.length === 0 || multiTypeConfig === null) {
			this.notify(output);
			return;
		}
		for(let typeConfig of multiTypeConfig) {
			let config: string[] = typeConfig.split(";");
			if(config[0] === tableType) {
				selectedTypeConfig = config;
				break;
			}
		}
		if(selectedTypeConfig.length === 0 || selectedTypeConfig.length === null) {
			this.notify(output);
			return;
		}
		let selectCols: string[] = selectedTypeConfig[1].split(",");
		this._context.webAPI.retrieveRecord(tableType, id,
			"?$select="+selectedTypeConfig[1]).then(
				(result) => {
					output = selectedTypeConfig !== undefined ? selectedTypeConfig[2]: "";
					if(output !== "") {
						selectCols.forEach(c=> {
							output = output.replace(c, result[c]);
							});
							this.notify(output);
					}
				},
			);
	}
}