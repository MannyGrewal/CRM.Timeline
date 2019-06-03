import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
const visTimeLine = require("vis/lib/timeline/Timeline");
const visDataSet = require("vis/lib/DataSet")
const moment = require("moment")
const RowRecordId:string = "rowRecId";

export class BasicTimeline implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Cached context object for the latest updateView
	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;
	private visuContainer: HTMLDivElement;
	private timeline : any;
	private _notifyOutputChanged: () => void;

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
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Need to track container resize so that control could get the available width. The available height won't be provided even this is true
			context.mode.trackContainerResize(true);
			this._notifyOutputChanged = notifyOutputChanged;

			// Create main table container div. 
			this.mainContainer = document.createElement("div");
			this.visuContainer = document.createElement("div");
			this.visuContainer.setAttribute("id","visualisation");
			this.mainContainer.appendChild(this.visuContainer);
			container.appendChild(this.mainContainer);

			

			

	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.contextObj = context;
		if(!context.parameters.dataSetGrid.loading){
				
			// Get sorted columns on View
			let columnsOnView = this.getSortedColumnsOnView(context);

			if (!columnsOnView || columnsOnView.length === 0) {
				return;
			}

			this.createTimeLine(columnsOnView,context.parameters.dataSetGrid);
		}
	}

	private createTimeLine(columnsOnView: DataSetInterfaces.Column[], ratings: DataSet){
		const dateColumn1 = 'scheduledstart';
		const dateColumn2 = 'scheduledend';
		const labelColumn = 'subject';

		if(ratings.sortedRecordIds.length > 0)
		{
			var timeLineItems = new visDataSet([]);
			let i=0;
			

			for(let currentRecordId of ratings.sortedRecordIds){
				let startdate = ratings.records[currentRecordId].getFormattedValue(dateColumn1);
				let enddate = ratings.records[currentRecordId].getFormattedValue(dateColumn2);
				let label = ratings.records[currentRecordId].getFormattedValue(labelColumn);

				if(startdate!= null && startdate != "" && label!= null && label != "")
				{
					let item = {
						id: ++i,
						content:label,
						start: moment(startdate,"D/MM/YYYY h:mm A").format('YYYY-MM-DD h:mm A'),
						end : (enddate!=null) ? moment(enddate,"D/MM/YYYY h:mm A").format('YYYY-MM-DD h:mm A') : "",
						type : this.DateDiff(enddate,startdate)<1 ? 'point' : 'box',
						title: ""
					}
					item["title"] = (item.end!="") ? item.start + "  -  " + item.end : item.start;
					timeLineItems._addItem(item); 
				}
			}

				// Configuration for the Timeline
			var options = {};

			// Create a Timeline
			if(!this.timeline)
				this.timeline = new visTimeLine(this.visuContainer, timeLineItems, options);			
		}
	}

	private DateDiff(end:any,start:any):number{
		if(end!=null && start!=null){
			return moment(end,"D/MM/YYYY h:mm A").diff(moment(start,"D/MM/YYYY h:mm A"),'days');
		}
		else
		return -1;
		
	}

		/**
		 * Get sorted columns on view
		 * @param context 
		 * @return sorted columns object on View
		 */
		private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[]
		{
			if (!context.parameters.dataSetGrid.columns) {
				return [];
			}
			
			let columns =context.parameters.dataSetGrid.columns
				.filter(function (columnItem:DataSetInterfaces.Column) { 
					// some column are supplementary and their order is not > 0
					return columnItem.order >= 0 }
				);
			
			// Sort those columns so that they will be rendered in order
			columns.sort(function (a:DataSetInterfaces.Column, b: DataSetInterfaces.Column) {
				return a.order - b.order;
			});
			
			return columns;
		}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}