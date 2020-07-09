import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Input, Loader, Button, Flex, FlexItem, Text, Icon as FluentIcon, Dropdown, DropdownProps, Checkbox, TextArea, Grid } from '@fluentui/react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { AxiosResponse } from "axios";
// import "./CreateIncident.scss";
import { isNullOrUndefined } from 'util';
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
import { orderBy, forEach } from 'lodash';
let reactPlugin = new ReactPlugin();
const browserHistory = createBrowserHistory({ basename: '' });


export interface IWorkstreamState {
    id: string,
    shortDescription: string,
    description: string,
    notes: string,
    isToggled: boolean,
    loader: boolean,
    status: string,
    workstreams: IWorkstream[],
    allBridges: IConferenceRooms[],
    selectedBridge: IConferenceRooms
}

export interface IWorkstream {
    priority: number,
    description: string,
    assignedTo: string,
    completed: boolean,
    assignedToId: string,
    inActive: boolean,
    id: string,
    partitionKey: string
}

export interface IConferenceRooms {
    code: number,
    available: boolean,
    bridgeURL: string,
    channelId: string,
}

export interface IIncident {
    bridge: string,
    description: string,
    impact: string,
    number: string,
    priority: string,
    short_description: string,
    state: string,
    sys_created_on: string,
    sys_id: string,
    sys_updated_on: string,
    bridgeDetails: IConferenceRooms
}

export const Priority = {
    Low: 1,
    Normal: 2,
    High: 3
}
const todayDate: Date = new Date();

export default class EditWorkstream extends React.Component<{}, IWorkstreamState> {

    token?: string | null = null;
    telemetry: any = undefined;
    incidentNumber?: string | null = null
    // appInsights: ApplicationInsights;

    constructor(props: {}) {
        super(props);
        initializeIcons();
        let startDate: Date = new Date();
        startDate.setHours(8, 30, 0, 0);
        let endDate: Date = new Date();
        endDate.setHours(9, 0, 0, 0);
        let workstream: IWorkstream = {
            priority: 1,
            description: "",
            assignedTo: "",
            completed: false,
            assignedToId: "",
            inActive: false,
            id: "",
            partitionKey: ""
        }
        this.state = {
            id: "",
            shortDescription: "",
            description: "",
            notes: "",
            loader: false,
            status: "New",
            isToggled: false,
            workstreams: [workstream],
            allBridges: [],
            selectedBridge: {
                available: true,
                channelId: "",
                code: 0,
                bridgeURL: ""
            }
        }
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
        this.incidentNumber = params.get("incident");
        // this.appInsights = new ApplicationInsights({
        //     config: {
        //         instrumentationKey: this.telemetry,
        //         extensions: [reactPlugin],
        //         extensionConfig: {
        //             [reactPlugin.identifier]: { history: browserHistory }
        //         }
        //     }
        // });
        // this.appInsights.loadAppInsights();

    };

    public componentDidMount = () => {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            console.log("microsoft teams", context)
        });
        document.removeEventListener("keydown", this.escFunction, false);
        this.getWorkstreams();
    }

    public componentWillUnmount = () => {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    private escFunction = (e: KeyboardEvent) => {
        if (e.keyCode === 27 || (e.key === "Escape")) {
            microsoftTeams.tasks.submitTask({ "output": "failure" });
        }
    }


    private getWorkstreams = async () => {
        this.setState({
            loader: true
        });
        await fetch("/api/WorkstreamApi/GetAllWorkstremsAsync?incidentNumber=" + this.incidentNumber, {
            method: "GET",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
        }).then(async (res) => {
            if (res.status === 401) {
                const response = await res.json();
                if (response) {
                    // this.setState({
                    //     errorResponseDetail: {
                    //         errorMessage: response.message,
                    //         statusCode: response.code,
                    //     }
                    // })
                }

                // this.setState({ authorized: false });
                // this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
                return response;
            }
            else if (res.status === 200) {
                let response = await res.json();
                // this.setState({
                //     loader: false
                // }, () => {
                let allWorkstreams: IWorkstream[] = [];
                for (let i = 0; i < response.length; i++) {
                    let workstream: IWorkstream = {
                        id: response[i].Id,
                        assignedTo: response[i].AssignedTo,
                        assignedToId: response[i].AssignedToId,
                        completed: response[i].Completed,
                        description: response[i].Description,
                        inActive: response[i].InActive,
                        partitionKey: response[i].PartitionKey,
                        priority: response[i].Priority,
                    }
                    allWorkstreams.push(workstream);
                }
                this.setState({
                    workstreams: allWorkstreams,
                    loader: false
                }, () => { console.log("LOG", this.state.workstreams) });
                // });
            }
            else {
                // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }

        });
    }

    private addWorkstreams = () => {
        let workstream: IWorkstream = {
            priority: this.state.workstreams.length + 1,
            description: "",
            assignedTo: "12345",
            completed: false,
            assignedToId: "123456",
            inActive: false,
            id: "",
            partitionKey: ""
        }
        this.setState({
            workstreams: [...this.state.workstreams, workstream],
        }, () => {
            // let priorityList = this.state.workstreams.map((workstream:IWorkstream, index: number)=>{
            //     console.log("List", index)
            //     return (index + 1);
            // });
            // this.list = priorityList;
        })
    }

    private onWorkstreamDescriptionChange = (e: React.SyntheticEvent<HTMLElement, Event>, index: number) => {
        let workstreams = this.state.workstreams;
        workstreams[index].description = (e.target as HTMLInputElement).value;
        this.setState({
            workstreams: workstreams
        })
    }

    private onWorkstreamAssigneeChange = (e: React.SyntheticEvent<HTMLElement, Event>, index: number) => {
        let workstreams = this.state.workstreams;
        workstreams[index].assignedTo = (e.target as HTMLInputElement).value;
        this.setState({
            workstreams: workstreams
        })
    }

    private setPriority = (e: React.SyntheticEvent<HTMLElement, Event>, dropdownProps?: DropdownProps) => {
        let workstreamSection = this.state.workstreams;
        let currentIndex: number = (Number)(dropdownProps!.defaultValue!);
        let newIndex: number = (Number)(dropdownProps!.value!);
        console.log("Priority", currentIndex, newIndex)
        workstreamSection[currentIndex - 1].priority = newIndex;
        workstreamSection[newIndex - 1].priority = currentIndex;
        workstreamSection.sort((a, b) => (a.priority > b.priority) ? 1 : ((b.priority > a.priority) ? -1 : 0));
        console.log(workstreamSection)
        this.setState({
            workstreams: workstreamSection
        });
    }

    private onDragOverItems = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
    }

    private onDragStartItems = (e: React.DragEvent<HTMLDivElement>, id: string) => {
        e.dataTransfer.setData("Id", id);
    }

    private onDropItems = (e: React.DragEvent<HTMLDivElement>, droppedOrder: number, droppedItemId: string) => {
        let draggedItemId: string = e.dataTransfer.getData("Id");

        let draggedItem = this.state.workstreams.filter(x => x.id === draggedItemId)[0];
        let droppedItem = this.state.workstreams.filter(x => x.id === droppedItemId)[0];

        let draggedIndex = this.state.workstreams.findIndex(x => x.id === draggedItemId)
        let droppedIndex = this.state.workstreams.findIndex(x => x.id === droppedItemId);

        let workstreams: IWorkstream[] = [];

        if (draggedIndex < droppedIndex) {
            for (let i = 0; i < this.state.workstreams.length; i++) {
                if (this.state.workstreams[i].id === draggedItem.id) {
                    let workstream: IWorkstream = this.state.workstreams[i];
                    workstream.priority = droppedOrder;
                    workstreams.push(workstream);
                }
                else if (this.state.workstreams[i].id === droppedItem.id) {
                    let workstream: IWorkstream = this.state.workstreams[i];
                    workstream.priority = droppedOrder - 1;
                    workstreams.push(workstream);
                }
                else {
                    if (i > draggedIndex && i < droppedIndex) {
                        let workstream: IWorkstream = this.state.workstreams[i];
                        workstream.priority = this.state.workstreams[i].priority - 1;
                        workstreams.push(workstream);
                    }
                    else {
                        let workstream: IWorkstream = this.state.workstreams[i];
                        workstreams.push(workstream);
                    }
                }
            }
        }
        else {
            for (let i = 0; i < this.state.workstreams.length; i++) {
                if (this.state.workstreams[i].id === draggedItem.id) {
                    let workstream: IWorkstream = this.state.workstreams[i];
                    workstream.priority = droppedOrder + 1;
                    workstreams.push(workstream);
                }
                else if (this.state.workstreams[i].id === droppedItem.id) {
                    let workstream: IWorkstream = this.state.workstreams[i];
                    workstream.priority = droppedOrder;
                    workstreams.push(workstream);
                }
                else {
                    if (i < draggedIndex && i > droppedIndex) {
                        let workstream: IWorkstream = this.state.workstreams[i];
                        workstream.priority = this.state.workstreams[i].priority + 1;
                        workstreams.push(workstream);
                    }
                    else {
                        let workstream: IWorkstream = this.state.workstreams[i];
                        workstreams.push(workstream);
                    }
                }
            }
        }
        workstreams = orderBy(workstreams, [item => item.priority]);

        this.setState({
            workstreams: workstreams,
        });
    }

    public render(): JSX.Element {
        const inputItems = [
            'New',
            'Suspended',
            'Service Restored'
        ];

        let workstreamBlock: JSX.Element[] = (this.state.workstreams.map((workstream: IWorkstream, index: number) => {
            console.log("Refresh!", workstream)
            let items = this.state.workstreams.map(item => item.priority)
            let description = <Text key={"description" + index} content={workstream.description} />
            if (workstream.description === ""){
                description = <Input className="inputField" defaultValue={workstream.assignedTo} value={workstream.assignedTo} key={"assignedto" + index} placeholder="Assigned to" fluid name="assignedto" onChange={(e) => this.onWorkstreamAssigneeChange(e, index)} />
            }
            return (
                <div draggable onDragOver={this.onDragOverItems} onDragStart={(e) => this.onDragStartItems(e, workstream.id)} onDrop={(e) => this.onDropItems(e, index, workstream.id)}>
                    <Flex gap="gap.small">
                        {/* <Dropdown
                            items={items}
                            defaultValue={items[index]}
                            value={items[index]}
                            placeholder="Start typing a name"
                            noResultsMessage="We couldn't find any matches."
                            onSelectedChange={this.setPriority}
                            key={"number" + index}
                        /> */}
                        <Text key={"number" + index} content={workstream.priority} />

                        {description}
                        <Text key={"assignedTo" + index} content={workstream.assignedTo} />
                        <Checkbox key={"completed" + index} label="Completed" checked={workstream.completed}/>
                        <Dropdown
                            search
                            items={inputItems}
                            placeholder="Type Text"
                            noResultsMessage="We couldn't find any matches."
                        />
                    </Flex>
                    <div hidden={index !== this.state.workstreams.length - 1}>
                        <Flex gap="gap.smaller">
                            <Icon iconName="add" className="pos-rel ft-18 ft-bld icon-sm" />
                            <Button text content="Add another workstream" onClick={this.addWorkstreams} />
                        </Flex>
                    </div>
                    <br />
                </div>
            )
        }));

        if (this.state.loader) {
            return (
                <div className="emptyContent">
                    <Loader />
                </div>
            );
        }
        else {
            return (
                <div className="taskModule">
                    <div className="formContainer">
                        <Grid>
                            {workstreamBlock}
                        </Grid>
                    </div>
                    <div className="footerContainer">
                        <div className="buttonContainer">
                            <Flex gap="gap.small">
                                <Button content="Submit" primary className="bottomButton" />
                            </Flex>
                        </div>
                    </div>
                </div >
            );
        }
    }
}
