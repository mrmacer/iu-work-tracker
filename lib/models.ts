export type SyncState = "saved" | "saving" | "sync-error";
export type WorkRecord = { appId:string; title:string; activityDate:string; activityType:string; description:string; detailedNotes:string; durationMinutes:number; status:"complete"|"draft"; projectIds:string[]; organizationIds:string[]; categoryIds:string[]; reach:{educatorsLeaders:number;studentsFamilies:number;workforceCommunity:number;other:number}; output:string; outcome:string; nextStep:string; followUpNeeded:boolean; followUpDate:string|null; orbit:{reportable:boolean;primaryDeliverable:string|null;supportingDeliverables:string[];stemPocMinutes:number;tacMinutes:number;evidence:string}; isSample:boolean; createdAt:string; modifiedAt:string; syncState?:SyncState };
export type Project = { appId:string;name:string;description:string;status:"active"|"planning"|"complete";color:string };
export type Organization = { appId:string;name:string;type:"district"|"partner"|"iu"|"regional" };
export type Category = { appId:string;name:string;group:"work-area"|"topic" };
export const PROJECTS:Project[]=[
  {appId:"project-steels",name:"STEELS Implementation",description:"District planning, professional learning, and implementation support.",status:"active",color:"blue"},
  {appId:"project-ai",name:"AI in Education",description:"Responsible AI learning and instructional support.",status:"active",color:"coral"},
  {appId:"project-keystone",name:"Keystone STEM Competition",description:"Regional student competition planning and delivery.",status:"planning",color:"lime"},
  {appId:"project-ecosystem",name:"STEM Ecosystem",description:"Cross-sector partnership and regional ecosystem development.",status:"active",color:"purple"},
  {appId:"project-makerspace",name:"Makerspace",description:"Hands-on student and educator experiences.",status:"active",color:"yellow"},
];
export const ORGANIZATIONS:Organization[]=[
  {appId:"org-north-valley",name:"North Valley SD",type:"district"},{appId:"org-riverbend",name:"Riverbend Area SD",type:"district"},
  {appId:"org-iu",name:"Intermediate Unit",type:"iu"},{appId:"org-futureworks",name:"FutureWorks Partnership",type:"partner"},{appId:"org-regional",name:"Regional / Multiple LEAs",type:"regional"},
];
export const CATEGORIES:Category[]=[
  {appId:"cat-district",name:"District Support",group:"work-area"},{appId:"cat-pl",name:"Professional Learning",group:"work-area"},{appId:"cat-stem",name:"STEM / Science",group:"topic"},
  {appId:"cat-steels",name:"STEELS",group:"topic"},{appId:"cat-ai",name:"Artificial Intelligence",group:"topic"},{appId:"cat-students",name:"Student Programs",group:"work-area"},
  {appId:"cat-partnerships",name:"Partnerships",group:"work-area"},{appId:"cat-internal",name:"Internal IU Work",group:"work-area"},{appId:"cat-planning",name:"Meetings / Planning",group:"work-area"},
];
export const DELIVERABLES=[["A","Statewide STEM & CS systems"],["B","PA STEELS implementation"],["C","CS, AI & computational thinking"],["D","Educational leadership"],["E","Workforce & ecosystem development"],["F","Student competitions & experiences"],["G","Math instruction & data literacy"]] as const;
export const MINUTES_PER_REPORTING_DAY=420;
