export interface Body {
    content: string;
    contentType: string;
}

export interface LinkedResource {
    webUrl: string;
    applicationName: string;
    displayName: string;
    externalId: string;
    id: string;
}

export interface Task {
    "@odata.etag": string;
    importance: string;
    isReminderOn: boolean;
    status: string;
    title: string;
    createdDateTime: string;
    lastModifiedDateTime: string;
    categories: any[];
    id: string;
    body: Body;
    "linkedResources@odata.context"?: string;
    linkedResources?: LinkedResource[];
}

export interface TasksCardData {
    cardTitle: string;
    tasks: Task[];
}