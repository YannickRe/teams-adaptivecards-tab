import { TodoTask } from "@microsoft/microsoft-graph-types";

export interface TasksCardData {
    cardTitle: string;
    tasks: TodoTask[];
}