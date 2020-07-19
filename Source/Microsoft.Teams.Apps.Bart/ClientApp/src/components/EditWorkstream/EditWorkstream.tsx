import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Input, Loader, Button, Flex, FlexItem, Text, Icon as FluentIcon, Dropdown, DropdownProps, Checkbox, CheckboxProps } from '@fluentui/react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Guid } from "guid-typescript";
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { AxiosResponse } from "axios";
import "./EditWorkstream.scss";
import { isNullOrUndefined } from 'util';
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
import { orderBy, forEach } from 'lodash';
let reactPlugin = new ReactPlugin();
const browserHistory = createBrowserHistory({ basename: '' });


export interface IWorkstreamState {
    id: string,
    isToggled: boolean,
    loader: boolean,
    status: string,
    workstreams: IWorkstream[],
    allBridges: IConferenceRooms[],
    selectedBridge: IConferenceRooms,
    users: IUser[],
    incidentAssignees: IUser[],
    incidentAssignedTo: IUser,
}

export interface IUser {
    id: string,
    displayName: string,
    userPrincipalName: string
}

export interface IWorkstream {
    priority: number,
    description: string,
    assignedTo: string,
    status: boolean,
    assignedToId: string,
    inActive: boolean,
    id: string,
    partitionKey: string,
    new: boolean,
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
    incidentNumber?: string | null = null;
    incidentId?: string | null = null;
    assignedToChanged?: boolean = false;
    // appInsights: ApplicationInsights;
    workstream: IWorkstream = {
        priority: 1,
        description: "",
        assignedTo: "",
        status: false,
        assignedToId: "",
        inActive: false,
        id: Guid.create().toString(),
        partitionKey: this.incidentNumber!,
        new: true
    }


    constructor(props: {}) {
        super(props);
        initializeIcons();
        let startDate: Date = new Date();
        startDate.setHours(8, 30, 0, 0);
        let endDate: Date = new Date();
        endDate.setHours(9, 0, 0, 0);
        this.state = {
            id: "",
            loader: false,
            status: "New",
            isToggled: false,
            workstreams: [this.workstream],
            allBridges: [],
            selectedBridge: {
                available: true,
                channelId: "",
                code: 0,
                bridgeURL: ""
            },
            users: [],
            incidentAssignees: [],
            incidentAssignedTo: {
                displayName: "null",
                id: "null",
                userPrincipalName: "null",
            }
        }
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
        this.incidentNumber = params.get("incident");
        this.incidentId = params.get("id");
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
        this.getAssignees();
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

        await Promise.all([
            fetch("/api/WorkstreamApi/GetAllWorkstremsAsync?incidentNumber=" + this.incidentNumber, {
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + this.token
                },
            }),
            fetch("/api/IncidentApi/AssignedUser?incidentNumber=" + this.incidentNumber, {
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + this.token
                },
            })
        ]).then(async (res) => {
            if (res[0].status === 200) {
                let response = await res[0].json();
                let allWorkstreams: IWorkstream[] = [];
                for (let i = 0; i < response.length; i++) {
                    let workstream: IWorkstream = {
                        id: response[i].Id,
                        assignedTo: response[i].AssignedTo,
                        assignedToId: response[i].AssignedToId,
                        status: response[i].Status,
                        description: response[i].Description,
                        inActive: response[i].InActive,
                        partitionKey: response[i].PartitionKey,
                        priority: response[i].Priority,
                        new: response[i].New
                    }
                    allWorkstreams.push(workstream);
                }
                this.workstream.priority = allWorkstreams.length + 1;
                this.workstream.partitionKey = this.incidentNumber!
                this.setState({
                    workstreams: [...orderBy(allWorkstreams, [items => items.priority]), this.workstream],
                    loader: false
                }, () => { console.log("LOG", this.state.workstreams) });
                // });

            }

            if (res[1].status === 200) {
                let response: IUser = await res[1].json();
                this.setState({
                    incidentAssignedTo: response
                })
            }
        });
    }

    private addWorkstreams = () => {
        let workstream: IWorkstream = {
            priority: this.state.workstreams.length + 1,
            description: "",
            assignedTo: "",
            status: false,
            assignedToId: "",
            inActive: false,
            id: Guid.create().toString(),
            partitionKey: this.incidentNumber!,
            new: true
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

    private getUsers = (e: React.SyntheticEvent<HTMLElement, Event>, data?: DropdownProps) => {
        var searchQuery = data!.searchQuery!
        this.searchUsers(searchQuery).then((res: any) => {
        });
    }

    private searchUsers = async (searchQuery: string) => {
        await fetch("/api/ResourcesApi/GetUsersAsync?fromFlag=1&searchQuery=" + searchQuery, {
            method: "POST",
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
            }
            else if (res.status === 200) {
                let response = await res.json();
                this.setState({
                    loader: false,
                    users: response
                });
                // let values: IUser[] = response.map((users: IUser)=>{
                //     let user: IUser = {
                //         displayName: users.displayName,
                //         id: users.id,
                //         userPrincipalName: users.userPrincipalName
                //     }
                // })
                // for (let i =0; i < response.length; i++){
                //     let user: IUser = {
                //         displayName: response[i].displayName,
                //         id: response[i].id,
                //         userPrincipalName: response[i].userPrincipalName
                //     }
                //     values.push(user);
                // }
            }
            else {
                // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }

        });
    }

    private getAssignees = async () => {
        await fetch("/api/ResourcesApi/GetUsersAsync?fromFlag=0", {
            method: "POST",
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
            }
            else if (res.status === 200) {
                let response = await res.json();
                this.setState({
                    loader: false,
                    incidentAssignees: response
                });
            }
            else {
                // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }

        });
    }

    private userAssigned = (e: React.SyntheticEvent<HTMLElement, Event>, v?: DropdownProps) => {
        let index = v! as { id: number }
        let selectedUser = v!.value! as { header: string, content: string };
        console.log("Users chanegs", selectedUser, index.id)
        var workstream = this.state.workstreams;
        workstream[index.id].assignedTo = selectedUser.header;
        workstream[index.id].assignedToId = this.state.users.find((user) => user.userPrincipalName === selectedUser.content)!.id!;

        this.setState({
            workstreams: workstream
        });

    }

    private assigneeChanged = (e: React.SyntheticEvent<HTMLElement, Event>, v?: DropdownProps) => {
        let currentUser = this.state.incidentAssignedTo;
        let selectedUser = v!.value! as { header: string, content: string };
        if (currentUser.id !== this.state.incidentAssignees.find((user) => user.userPrincipalName === selectedUser.content)!.id!) {
            var user: IUser = {
                displayName: selectedUser.header,
                id: this.state.incidentAssignees.find((user) => user.userPrincipalName === selectedUser.content)!.id!,
                userPrincipalName: ""
            };
            console.log("Assign chanegs", selectedUser, currentUser)

            this.setState({
                incidentAssignedTo: user
            });
            this.assignedToChanged = true;
        }
    }

    private onWorkstreamAssigneeChange = (e: React.SyntheticEvent<HTMLElement, Event>, index: number) => {
        let workstreams = this.state.workstreams;
        workstreams[index].assignedTo = (e.target as HTMLInputElement).value;
        this.setState({
            workstreams: workstreams
        })
    }

    private onDragOverItems = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
    }

    private onDragStartItems = (e: React.DragEvent<HTMLDivElement>, id: string) => {
        e.dataTransfer.setData("Id", id);
    }

    private onDropItems = (e: React.DragEvent<HTMLDivElement>, droppedOrder: number, droppedItemId: string) => {
        console.log("drop-droppedOrder", droppedOrder, droppedItemId)

        let draggedItemId: string = e.dataTransfer.getData("Id");

        let draggedItem = this.state.workstreams.filter(x => x.id === draggedItemId)[0];
        let droppedItem = this.state.workstreams.filter(x => x.id === droppedItemId)[0];
        if (droppedItem.description === "") {
            e.preventDefault();
            return;
        }

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

    private completedCheckboxChanged = (e: React.SyntheticEvent<HTMLElement, Event>, v?: CheckboxProps) => {
        const selectedId = (e.currentTarget as Element).id;
        this.state.workstreams.forEach((currentItem) => {
            if (currentItem.id === selectedId) {
                currentItem.status = v!.checked ? v!.checked : false;
            }
        });
        this.setState({
            workstreams: this.state.workstreams
        });
    }

    private removeWorkstream = (workstreamId: string) => {
        let allWorkstrems = this.state.workstreams;
        allWorkstrems.forEach((currentWorkstream) => {
            if (currentWorkstream.id === workstreamId) {
                currentWorkstream.inActive = true;
                allWorkstrems = allWorkstrems.map((workstream) => {
                    if (workstream.priority > currentWorkstream.priority) {
                        workstream.priority--;
                        return workstream;
                    } else {
                        return workstream;
                    }
                })
            }
        });
        this.setState({
            workstreams: allWorkstrems
        });
    }

    private submitWorkstreams = async () => {
        this.setState({
            loader: true
        })

        let assignedToObject = {
            AssignedTo: this.state.incidentAssignedTo.displayName,
            AssignedToId: this.state.incidentAssignedTo.id,
            PartitionKey: this.incidentNumber,
            RowKey: this.incidentId
        }

        let requests = [
            fetch("/api/WorkstreamApi/CreateOrUpdateWorkstremAsync", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + this.token
                },
                body: JSON.stringify(this.state.workstreams)
            })
        ];
        if (this.assignedToChanged) {
            requests.push(
                fetch("/api/IncidentApi/AssignTicket", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "Authorization": "Bearer " + this.token
                    },
                    body: JSON.stringify(assignedToObject)
                })
            )
        }
        await Promise.all(requests).then(async (res) => {

            let assignedToObject = {
                assignedToId: this.state.incidentAssignedTo.id,
                assignedTo: this.state.incidentAssignedTo.displayName,
                incidentNumber: this.incidentNumber,
                output: this.assignedToChanged
            }

            microsoftTeams.tasks.submitTask(assignedToObject);
        });
    }

    public render(): JSX.Element {
        const inputItems = [
            'New',
            'Suspended',
            'Service Restored'
        ];

        const userDetails = this.state.users.map((user) => {
            console.log("UserDetails", user)
            return ({
                header: user.displayName,
                content: user.userPrincipalName
            });
        });

        const incidentAssignees = this.state.incidentAssignees.map((user) => {
            console.log("incidentAssignees", user)
            return ({
                header: user.displayName,
                content: user.userPrincipalName
            });
        });

        let workstreamBlock = (this.state.workstreams.map((workstream: IWorkstream, index: number) => {
            if (!workstream.inActive) {
                console.log("Refresh!", workstream)
                let description = <Text key={"description" + index} content={workstream.description} />
                let assignedTo = <Text key={"assignedTo" + index} content={workstream.assignedTo} />

                if (workstream.description === "" || this.state.workstreams.length - 1 === index) {
                    description = <Input className="inputField" defaultValue={workstream.description} value={workstream.description} key={"assignedto" + index} placeholder="Description" fluid name="assignedto" onChange={(e) => this.onWorkstreamDescriptionChange(e, index)} />
                }

                if (workstream.assignedTo === "" || this.state.workstreams.length - 1 === index) {
                    assignedTo = <Dropdown
                        search
                        clearable
                        id={index.toString()}
                        items={userDetails}
                        defaultValue={workstream.assignedTo}
                        value={workstream.assignedTo}
                        placeholder="Assign to"
                        onSearchQueryChange={this.getUsers}
                        noResultsMessage="We couldn't find any matches."
                        onSelectedChange={this.userAssigned}
                        key={"number" + index}
                    />
                }

                return (
                    <tr draggable={workstream.description !== ""} onDragOver={this.onDragOverItems} onDragStart={(e) => this.onDragStartItems(e, workstream.id)} onDrop={(e) => this.onDropItems(e, index + 1, workstream.id)}>
                        <td>
                            <Text key={"number" + index} content={workstream.priority} />
                        </td>
                        <td>
                            {description}
                        </td>
                        <td>
                            {assignedTo}
                        </td>
                        <td>
                            <Checkbox key={"completed" + index} id={workstream.id} label="Completed" checked={workstream.status} onClick={this.completedCheckboxChanged} />
                        </td>
                        <td>
                            <Button className="close-btn" content={<Icon iconName="trash" className="deleteIcon" />} iconOnly title="Close"
                                onClick={() => this.removeWorkstream(workstream.id)} />
                        </td>

                        {/* <div hidden={index !== this.state.workstreams.length - 1}>
                            <Flex gap="gap.smaller">
                                <Icon iconName="add" className="addIcon" />
                                <Button text content="Add another workstream" onClick={this.addWorkstreams} />
                            </Flex>
                        </div> */}
                    </tr>
                )
            }
        }));

        if (this.state.loader) {
            return (
                <div className="emptyContent">
                    <Loader />
                </div>
            );
        }
        else {
            console.log("workstreamBlock", workstreamBlock)
            return (
                <div className="taskModule">
                    <div className="formContainer ">

                        <Dropdown

                            items={incidentAssignees}
                            placeholder="Assign Incident"
                            noResultsMessage="We couldn't find any matches."
                            onSelectedChange={this.assigneeChanged}
                            defaultValue={isNullOrUndefined(this.state.incidentAssignedTo.id) ? "" :
                                this.state.incidentAssignedTo.displayName}
                        />
                        <h4>Here are few workstreams</h4>
                        <table className="table table-borderless">
                            <thead>
                                <tr>
                                    <th>Priority</th>
                                    <th>Description</th>
                                    <th>Assigned to</th>
                                    <th>Status</th>
                                    <th></th>
                                </tr>
                            </thead>
                            <tbody>
                                {workstreamBlock}
                            </tbody>
                        </table>
                        {/* <div hidden={index !== this.state.workstreams.length - 1}> */}
                        <Flex gap="gap.smaller">
                            <Icon iconName="add" className="addIcon" />
                            <Button text content="Add another workstream" onClick={this.addWorkstreams}
                                disabled={this.state.workstreams[this.state.workstreams.length - 1].description === ""
                                    && this.state.workstreams[this.state.workstreams.length - 1].assignedTo === ""} />
                        </Flex>
                        {/* </div> */}
                    </div>
                    <div className="footerContainer">
                        <div className="buttonContainer">
                            <Flex gap="gap.small">
                                <FlexItem push>
                                    <Button content="Submit" primary className="bottomButton" onClick={this.submitWorkstreams} />
                                </FlexItem>
                            </Flex>
                        </div>
                    </div>
                </div >
            );
        }
    }
}