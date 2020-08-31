import { TemplateContext,ComponentMediaQuery,IDynamicPerson,PersonType,GroupType,PersonCardInteraction,PersonViewType,AvatarSize,TasksStringResource,TaskFilter,SelectedChannel } from '@microsoft/mgt';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type AgendaProps = {
	date?: string;
	groupId?: string;
	days?: number;
	eventQuery?: string;
	events?: MicrosoftGraph.Event[];
	showMax?: number;
	groupByDay?: boolean;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	eventClick?: (e: Event) => void;
	updated?: (e: Event) => void;
}

export type GetProps = {
	resource?: string;
	scopes?: string[];
	version?: string;
	maxPages?: number;
	pollingRate?: number;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	dataChange?: (e: Event) => void;
	updated?: (e: Event) => void;
}

export type LoginProps = {
	userDetails?: IDynamicPerson;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	loginInitiated?: (e: Event) => void;
	loginCompleted?: (e: Event) => void;
	loginFailed?: (e: Event) => void;
	logoutInitiated?: (e: Event) => void;
	logoutCompleted?: (e: Event) => void;
	updated?: (e: Event) => void;
}

export type PeoplePickerProps = {
	groupId?: string;
	type?: PersonType;
	groupType?: GroupType;
	people?: IDynamicPerson[];
	defaultSelectedUserIds?: string[];
	placeholder?: string;
	selectionMode?: string;
	showMax?: number;
	selectedPeople?: IDynamicPerson[];
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
	updated?: (e: Event) => void;
}

export type PeopleProps = {
	groupId?: string;
	userIds?: string[];
	people?: IDynamicPerson[];
	peopleQueries?: string[];
	showPresence?: boolean;
	personCardInteraction?: PersonCardInteraction;
	showMax?: number;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	updated?: (e: Event) => void;
}

export type PersonCardProps = {
	personQuery?: string;
	userId?: string;
	personDetails?: IDynamicPerson;
	personImage?: string;
	fetchImage?: boolean;
	isExpanded?: boolean;
	inheritDetails?: boolean;
	showPresence?: boolean;
	personPresence?: MicrosoftGraphBeta.Presence;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	updated?: (e: Event) => void;
}

export type PersonProps = {
	personQuery?: string;
	userId?: string;
	showName?: boolean;
	showEmail?: boolean;
	showPresence?: boolean;
	personDetails?: IDynamicPerson;
	personImage?: string;
	fetchImage?: boolean;
	personPresence?: MicrosoftGraphBeta.Presence;
	personCardInteraction?: PersonCardInteraction;
	line1Property?: string;
	line2Property?: string;
	view?: PersonViewType;
	avatarSize?: AvatarSize;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	updated?: (e: Event) => void;
}

export type TasksProps = {
	res?: TasksStringResource;
	isNewTaskVisible?: boolean;
	readOnly?: boolean;
	dataSource?: string;
	targetId?: string;
	targetBucketId?: string;
	initialId?: string;
	initialBucketId?: string;
	hideHeader?: boolean;
	hideOptions?: boolean;
	groupId?: string;
	taskFilter?: TaskFilter;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	taskAdded?: (e: Event) => void;
	taskChanged?: (e: Event) => void;
	taskClick?: (e: Event) => void;
	taskRemoved?: (e: Event) => void;
	updated?: (e: Event) => void;
}

export type TeamsChannelPickerProps = {
	selectedItem?: SelectedChannel;
	templateContext?: TemplateContext;
	templateConverters?: TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
	updated?: (e: Event) => void;
}

export const Agenda = wrapMgt<AgendaProps>('mgt-agenda');

export const Get = wrapMgt<GetProps>('mgt-get');

export const Login = wrapMgt<LoginProps>('mgt-login');

export const PeoplePicker = wrapMgt<PeoplePickerProps>('mgt-people-picker');

export const People = wrapMgt<PeopleProps>('mgt-people');

export const PersonCard = wrapMgt<PersonCardProps>('mgt-person-card');

export const Person = wrapMgt<PersonProps>('mgt-person');

export const Tasks = wrapMgt<TasksProps>('mgt-tasks');

export const TeamsChannelPicker = wrapMgt<TeamsChannelPickerProps>('mgt-teams-channel-picker');

