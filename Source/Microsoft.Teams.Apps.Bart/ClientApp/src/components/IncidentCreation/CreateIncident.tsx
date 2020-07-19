import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Input, Loader, Button, Flex, FlexItem, Text, Icon as FluentIcon, Dropdown, DropdownProps, Checkbox, TextArea } from '@fluentui/react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { AxiosResponse } from "axios";
import "./CreateIncident.scss";
import "./bootstrap-grid.css";
import { isNullOrUndefined } from 'util';
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
let reactPlugin = new ReactPlugin();
const browserHistory = createBrowserHistory({ basename: '' });

export interface ICreateIncidentProps {

}

export interface ICreateIncidentState {
    id: string,
    shortDescription: string,
    description: string,
    scope: string,
    isToggled: boolean,
    loader: boolean,
    status: string,
    workstreams: IWorkstream[],
    allBridges: IConferenceRooms[],
    selectedBridge: IConferenceRooms,
    users: IUser[]
}

export interface IWorkstream {
    priority: number,
    description: string,
    assignedTo: string,
    completed: boolean,
    assignedToId: string,
    inActive: boolean,
    new: boolean
}

export interface IUser {
    id: string,
    displayName: string,
    userPrincipalName: string
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
    bridgeDetails: IConferenceRooms,
}

export const Priority = {
    Low: 1,
    Normal: 2,
    High: 3
}
const todayDate: Date = new Date();

export default class CreateIncident extends React.Component<ICreateIncidentProps, ICreateIncidentState> {

    private shortDescription = "";
    private description = "";
    private scope = "";
    private status = "";
    private priority = 0;
    private list: number[] = [];
    token?: string | null = null;
    telemetry: any = undefined;
    fetchedDescription: string | null = null;
    requestedBy?: IUser | null = null;
    requestedFor?: IUser | null = null;
    // appInsights: ApplicationInsights;

    constructor(props: ICreateIncidentProps) {
        super(props);
        initializeIcons();
        let startDate: Date = new Date();
        startDate.setHours(8, 30, 0, 0);
        let endDate: Date = new Date();
        endDate.setHours(9, 0, 0, 0);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
        this.fetchedDescription = params.get("description");
        this.requestedBy = {
            displayName: params.get("displayName")!,
            id: "",
            userPrincipalName: ""
        };
        let workstream: IWorkstream = {
            priority: 1,
            description: "",
            assignedTo: "",
            completed: false,
            assignedToId: "",
            inActive: false,
            new: true
        }
        this.state = {
            id: "",
            shortDescription: "",
            description: this.fetchedDescription !== null ? this.fetchedDescription : "",
            scope: "",
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
            },
            users: []
        }
        this.shortDescription = "";
        this.description = this.fetchedDescription !== null ? this.fetchedDescription : "";
        this.scope = "";
        this.status = "";
        this.priority = 0;
        this.list = [1];
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
            this.requestedBy!.id = context.userObjectId!;
            this.requestedBy!.userPrincipalName = context.userPrincipalName!;
            console.log("microsoft teams", context, this.requestedBy!)
        });

        document.removeEventListener("keydown", this.escFunction, false);
        this.getAvailableBridges();
    }


    public componentWillUnmount = () => {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    private getAvailableBridges = async () => {
        this.setState({
            loader: true
        });
        await fetch("/api/ResourcesApi/GetAvailabilityData", {
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
                let bridges: IConferenceRooms[] = [];
                for (let i = 0; i < response.length; i++) {
                    bridges.push({
                        available: response[i].Available,
                        channelId: response[i].ChannelId,
                        code: response[i].Code,
                        bridgeURL: response[i].BridgeURL
                    })
                }
                this.setState({
                    loader: false,
                    allBridges: bridges,
                    selectedBridge: bridges.find((bridge) => bridge.code.toString() === "0")!
                }, () => {
                    console.log("=>", this.state.allBridges)
                });
            }
            else {
                // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }

        });
    }

    private escFunction = (e: KeyboardEvent) => {
        if (e.keyCode === 27 || (e.key === "Escape")) {
            microsoftTeams.tasks.submitTask({ "output": "failure" });
        }
    }

    private onShortDescriptionChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        console.log("SD", (e.target as HTMLInputElement).value)
        this.shortDescription = (e.target as HTMLInputElement).value;
    }

    private onDescriptionChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        console.log("D", (e.target as HTMLInputElement).value)
        this.description = (e.target as HTMLInputElement).value;
    }

    private onScopeChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        this.scope = (e.target as HTMLInputElement).value;
    }

    private onPriorityChange = (e: React.SyntheticEvent<HTMLElement, Event>, dropdownProps?: DropdownProps) => {
        console.log("Priority", Object.keys(Priority).map(key => { if (key === "Low") return Priority[key] })[0]!)
        // let priorityChoice: string = (String)(dropdownProps!.value!);
        // // const keys = Object.keys(Priority) as (keyof (Priority)[];
        // this.priority = Object.keys(Priority).map(key => 
        //     {
        //         if(key === priorityChoice) 
        //         return Priority[key];
        //     })[0]!;
    }

    // private onStatusChange = (e: React.SyntheticEvent<HTMLElement, Event>) =>{
    //     this.shortDescription = (e.target as HTMLInputElement).value;
    // }

    private getUsers = (e: React.SyntheticEvent<HTMLElement, Event>, data?: DropdownProps) => {
        var searchQuery = data!.searchQuery!
        this.searchUsers(searchQuery).then((res: any) => {
            // console.log("Users", res)

            // this.setState({
            //     users: res
            // },()=>{
            //     console.log("Users", this.state.users)
            // });
        });
    }

    private userAssigned = (e: React.SyntheticEvent<HTMLElement, Event>, v?: DropdownProps) => {
        let selectedUser = v!.value! as { header: string, content: string };
        let index = v! as { id: number }
        console.log("Users chanegs", selectedUser, index.id)
        var workstream = this.state.workstreams;
        workstream[index.id].assignedTo = selectedUser.header;
        workstream[index.id].assignedToId = this.state.users.find((user) => user.userPrincipalName === selectedUser.content)!.id!;;

        this.setState({
            workstreams: workstream
        });

    }

    private requestedAssigned = (e: React.SyntheticEvent<HTMLElement, Event>, v?: DropdownProps) => {
        let selectedUser = v!.value! as { header: string, content: string };
        let index = v! as { id: number }
        console.log("Users chanegs", selectedUser, index.id)
        if (selectedUser.content !== this.requestedBy!.userPrincipalName) {
            this.requestedFor = {
                id: this.state.users.find((user) => user.userPrincipalName === selectedUser.content)!.id!,
                displayName: selectedUser.header,
                userPrincipalName: selectedUser.content
            }
        }

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

    private createIncident = async () => {
        console.log("CreateIncident", this.description);
        this.setState({
            loader: true
        });
        let event = {
            Incident: {
                Short_Description: this.shortDescription,
                Description: this.description,
                Priority: 7,
                Bridge: this.state.selectedBridge.code,
                bridgeDetails: this.state.selectedBridge,
                Scope: this.scope
            },
            Workstreams: this.state.workstreams
        };
        await fetch("/api/IncidentApi/CreateIncidentAsync", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
            body: JSON.stringify(event)
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
                let toBot: IIncident = response;
                toBot.bridge = this.state.selectedBridge.code.toString();
                toBot.bridgeDetails = this.state.selectedBridge;
                microsoftTeams.tasks.submitTask(toBot);
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
            assignedTo: "",
            completed: false,
            assignedToId: "",
            inActive: false,
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

    private setBridge = (e: React.SyntheticEvent<HTMLElement, Event>, dropdownProps?: DropdownProps) => {
        console.log("setBridge", dropdownProps)
        let selectedBridge = this.state.allBridges.find(bridge => bridge.code === dropdownProps!.value!)
        this.setState({
            selectedBridge: selectedBridge!
        });
    }

    public render(): JSX.Element {
        const inputItems = [
            'New',
            'Suspended',
            'Service Restored'
        ];

        const userInput = this.state.users.map((user) => {
            console.log("UserDetails", user)
            return ({
                header: user.displayName,
                content: user.userPrincipalName
            });
        });

        console.log("User", userInput)

        let workstreamBlock: JSX.Element[] = (this.state.workstreams.map((workstream: IWorkstream, index: number) => {
            console.log("Refresh!", this.state.workstreams[index].description, workstream.description)
            let items = this.state.workstreams.map(item => item.priority)
            return (
                <div>
                    <div className="row my-1">
                        <div className="col-md-2 pr-1">
                            <Dropdown
                                className="xs-small-input"
                                items={items}
                                defaultValue={items[index]}
                                value={items[index]}
                                onSelectedChange={this.setPriority}
                                key={"number" + index}
                            />
                        </div>
                        <div className="col-md-5 px-1">
                            <Input className="inputField" defaultValue={workstream.description} value={workstream.description} key={"description" + index} placeholder="Description" fluid name="description" onChange={(e) => this.onWorkstreamDescriptionChange(e, index)} />
                        </div>
                        <div className="col-md-3 px-1">
                            {/* <Input className="inputField" defaultValue={workstream.assignedTo} value={workstream.assignedTo} key={"assignedto" + index} placeholder="Assigned to" fluid name="assignedto" onChange={(e) => this.onWorkstreamAssigneeChange(e, index)} /> */}
                            <Dropdown
                                className="md-input"
                                clearable
                                search
                                id={index.toString()}
                                onSearchQueryChange={this.getUsers}
                                items={userInput}
                                placeholder="Start typing a name"
                                onSelectedChange={this.userAssigned}
                                noResultsMessage="We couldn't find any matches."
                                value={workstream.assignedTo}
                            />
                        </div>
                        <div className="col-md-2 pl-1">
                            <Dropdown
                                className="md-input"
                                search
                                items={inputItems}
                                placeholder="Type Text"
                                noResultsMessage="We couldn't find any matches."
                            />
                        </div>
                    </div>
                    <div hidden={index !== this.state.workstreams.length - 1}>
                        <Flex gap="gap.smaller">
                            <Icon iconName="add" className="pos-rel ft-18 ft-bld icon-sm" />
                            <Button text content="Add another workstream" onClick={this.addWorkstreams}
                                disabled={this.state.workstreams[this.state.workstreams.length - 1].description === ""
                                    && this.state.workstreams[this.state.workstreams.length - 1].assignedTo === ""} />
                        </Flex>
                    </div>
                    <br />
                </div>
            )
        }));

        let bridgeCodes = this.state.allBridges.map(bridge => { if (bridge.code.toString() !== "0") return bridge.code });
        console.log("bridgeCodes", bridgeCodes)
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
                        <div className="row">
                            <div className="col-md-12">
                                <Flex gap="gap.smaller">
                                    <Text content="Individual requesting incident" />
                                </Flex>
                            </div>
                        </div>
                        <div className="custom">
                            <div className="row">
                                <div className="col-md-4 col-lg-4">
                                    <Dropdown
                                        className="md-input"
                                        clearable
                                        search
                                        id="requestedBy"
                                        onSearchQueryChange={this.getUsers}
                                        items={userInput}
                                        placeholder="Start typing a name"
                                        onSelectedChange={this.requestedAssigned}
                                        defaultValue={this.requestedBy!.displayName}
                                    />
                                </div>
                            </div>
                            <div className="row my-3">
                                <div className="col-md-8">
                                    <Flex gap="gap.smaller">
                                        <Text content="Short description(Note: max 250 characters)" />
                                    </Flex>
                                    <Flex gap="gap.smaller">
                                        <Input fluid className="inputField" defaultValue={this.shortDescription} placeholder="Short description" name="shortDescriptionTitle" onChange={this.onShortDescriptionChange} />
                                    </Flex>
                                </div>
                                <div className="col-md-4">
                                    <Flex gap="gap.smaller">
                                        <Text content="Scope" />
                                    </Flex>
                                    <Flex gap="gap.smaller">
                                        <Input fluid className="inputField" defaultValue={this.scope} placeholder="Scope" name="scopeTitle" onChange={this.onScopeChange} />

                                    </Flex>
                                </div>
                            </div>
                            <div className="row my-3">
                                <div className="col-md-8">
                                    <Flex gap="gap.smaller" column>
                                        <Text content="Description of the reported problem" />
                                        <TextArea fluid className="inputField textarea" defaultValue={this.description} placeholder="Description" name="descriptionTitle" onChange={this.onDescriptionChange} />
                                    </Flex>
                                </div>
                                <div className="col-md-4">

                                    <Flex gap="gap.smaller" column>
                                        <Text content="Conference bridge" />
                                        <Dropdown
                                            className="select-wrapper"
                                            items={bridgeCodes}
                                            placeholder="Select conference bridges"
                                            noResultsMessage="We couldn't find any matches."
                                            onSelectedChange={this.setBridge}
                                        />
                                    </Flex>
                                    <Flex gap="gap.smaller" column className="mt-3">
                                        <Text content="Priority" />
                                        <Dropdown
                                            className="select-wrapper"
                                            items={Object.keys(Priority)}
                                            placeholder="Select priority"
                                            noResultsMessage="We couldn't find any matches."
                                            onSelectedChange={this.onPriorityChange}
                                        />
                                        {/* <Input className="inputField" value={this.state.title} placeholder="Search title goes here" name="txtTitle" /> */}
                                    </Flex>
                                </div>
                            </div>
                            <div className="row my-3">
                                <div className="col-md-12">
                                    <Flex>
                                        <Text content="TSC request" className="mt-1 mr-2" />
                                        <Checkbox label="Did this request originated from technology support center" />
                                    </Flex>
                                </div>
                            </div>
                        </div>
                        <div className="row my-3">
                            <div className="col-md-12">
                                <Text className="h5 bold" content="Create workstream" />
                            </div>
                        </div>
                        <div className="row my-1">
                            <div className="col-md-2 pr-1">
                                <Text content="Priorty" />
                            </div>
                            <div className="col-md-5 px-1">
                                <Text content="Description" />
                            </div>
                            <div className="col-md-3 px-1">
                                <Text content="Assigned to" />
                            </div>
                            <div className="col-md-2 pl-1">
                                <Text content="Status" />
                            </div>
                        </div>
                        {workstreamBlock}

                        <div className="footerContainer">
                            <div className="buttonContainer">
                                <Flex gap="gap.small">
                                    <Button content="Submit" primary className="bottomButton" onClick={this.createIncident} />
                                </Flex>
                            </div>
                        </div>
                    </div>
                </div >
            );
        }
    }
}
