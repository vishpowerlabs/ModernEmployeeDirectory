export interface IKudos {
    id: string;
    recipientId: string;
    recipientName: string;
    authorId: string;
    authorName: string;
    message: string;
    badgeType: string;
    createdDate: Date;
}

export interface IKudosListItem {
    Id: number;
    [key: string]: any;
}
