import { IPropertyFieldPeople } from 'sp-client-custom-fields/lib/PropertyFieldPeoplePicker';

export interface IIssuesViewWebPartProps {
  description : string;
  people : IPropertyFieldPeople[];
  showForTeam: boolean;
  teamName: string;
}
