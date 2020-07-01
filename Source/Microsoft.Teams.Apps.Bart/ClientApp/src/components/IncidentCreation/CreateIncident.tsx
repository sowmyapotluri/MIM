import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Input, Loader, Button, Flex, FlexItem, Text, Icon as FluentIcon, Dropdown, DropdownProps, Checkbox, TextArea } from '@fluentui/react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { AxiosResponse } from "axios";
import "./CreateIncident.scss";
import { isNullOrUndefined } from 'util';
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
let reactPlugin = new ReactPlugin();
const browserHistory = createBrowserHistory({ basename: '' });


const DayPickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

    goToToday: '',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker'
};

const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px'
    }
});

export interface ICreateSessionProps {

}

export interface ICreateSessionState {
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
}

export interface IConferenceRooms {
    code: number,
    available: boolean,
    bridgeURL: string,
    channelId: string,
}

export interface IIncident {
    bridge: IConferenceRooms,
    description: string,
    impact: string,
    number: string,
    priority: string,
    short_description: string,
    state: string,
    sys_created_on: string,
    sys_id: string,
    sys_updated_on: string,
}

export const Priority = {
    Low: 1,
    Normal: 2,
    High: 3
}
const todayDate: Date = new Date();

export default class CreateIncident extends React.Component<ICreateSessionProps, ICreateSessionState> {

    private readonly newSessionStatus = "New";
    private readonly type = "Session";
    private shortDescription = "";
    private description = "";
    private notes = "";
    private status = "";
    private priority = 0;
    private list: number[] = [];
    token?: string | null = null;
    telemetry: any = undefined;
    // appInsights: ApplicationInsights;

    constructor(props: ICreateSessionProps) {
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
        this.shortDescription = "";
        this.description = "";
        this.notes = "";
        this.status = "";
        this.priority = 0;
        this.list = [1];
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
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
                    allBridges: bridges
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

    private onNotesChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        this.notes = (e.target as HTMLInputElement).value;
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

    private createIncident = async () => {
        this.setState({
            loader: true
        });
        let event = {
            Short_Description: this.shortDescription,
            Description: this.description,
            Priority: 7,
            Bridge: this.state.selectedBridge
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
                    toBot.bridge = this.state.selectedBridge;
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

        let workstreamBlock: JSX.Element[] = (this.state.workstreams.map((workstream: IWorkstream, index: number) => {
            console.log("Refresh!", this.state.workstreams[index].description, workstream.description)
            let items = this.state.workstreams.map(item => item.priority)
            return (
                <div>
                    <Flex gap="gap.small">
                        <Dropdown
                            items={items}
                            defaultValue={items[index]}
                            value={items[index]}
                            placeholder="Start typing a name"
                            noResultsMessage="We couldn't find any matches."
                            onSelectedChange={this.setPriority}
                            key={"number" + index}
                        />
                        <Input className="inputField" defaultValue={workstream.description} value={workstream.description} key={"description" + index} placeholder="Description" fluid name="description" onChange={(e) => this.onWorkstreamDescriptionChange(e, index)} />
                        <Input className="inputField" defaultValue={workstream.assignedTo} value={workstream.assignedTo} key={"assignedto" + index} placeholder="Assigned to" fluid name="assignedto" onChange={(e) => this.onWorkstreamAssigneeChange(e, index)} />
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

        let bridgeCodes = this.state.allBridges.map(bridge => bridge.code);
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

                        <Flex gap="gap.smaller">
                            <Text content="Individual requesting incident" />
                        </Flex>
                        <div className="custom">
                            <Flex gap="gap.smaller">
                                <Input className="inputField" value={this.state.id} name="txtPasskey" />
                            </Flex>
                            <br />
                            <Flex gap="gap.smaller">
                                <Text content="Short description(Note: max 250 characters)" />
                            </Flex>
                            <Flex gap="gap.smaller">
                                <Input fluid className="inputField" defaultValue={this.shortDescription} placeholder="Search title goes here" name="shortDescriptionTitle" onChange={this.onShortDescriptionChange} />
                            </Flex>
                            <br />
                            <Flex gap="gap.large">
                                <FlexItem grow>
                                    <Flex gap="gap.smaller" column>
                                        <Text content="Notes" />
                                        <Input fluid className="inputField" defaultValue={this.notes} placeholder="Search title goes here" name="notesTitle" onChange={this.onNotesChange} />
                                    </Flex>
                                </FlexItem>
                                <FlexItem push>
                                    <Flex gap="gap.smaller" column>
                                        <Text content="Status" />
                                        <Dropdown
                                            search
                                            items={inputItems}
                                            placeholder="Type Text"
                                            noResultsMessage="We couldn't find any matches."
                                        />
                                    </Flex>
                                </FlexItem>
                            </Flex>
                            <br />
                            <Flex gap="gap.smaller">
                                <FlexItem>
                                    <Flex gap="gap.smaller" column>
                                        <Text content="Description of the reported problem" />
                                        <TextArea fluid className="inputField" defaultValue={this.description} placeholder="Search title goes here" name="descriptionTitle" onChange={this.onDescriptionChange} />
                                        <Flex column>
                                            <Text content="TSC request" />
                                            <Checkbox label="Did this request originated from technology support center" />
                                        </Flex>
                                    </Flex>
                                </FlexItem>
                                <FlexItem>
                                    <Flex column>
                                        <Flex gap="gap.smaller" column>
                                            <Text content="Conference bridge" />
                                            <Dropdown
                                                items={bridgeCodes}
                                                placeholder="Type Text"
                                                noResultsMessage="We couldn't find any matches."
                                                onSelectedChange={this.setBridge}
                                            />
                                        </Flex>
                                        <Flex gap="gap.smaller" column>
                                            <Text content="Priority" />
                                            <Dropdown
                                                items={Object.keys(Priority)}
                                                placeholder="Type Text"
                                                noResultsMessage="We couldn't find any matches."
                                                onSelectedChange={this.onPriorityChange}
                                            />
                                            {/* <Input className="inputField" value={this.state.title} placeholder="Search title goes here" name="txtTitle" /> */}
                                        </Flex>
                                    </Flex>
                                </FlexItem>
                            </Flex>
                        </div>
                        <Text size="large" content="Create workstream" />
                        <br />
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
