import * as React from 'react';
import { Avatar, Chip } from '@mui/material';
import type { IPersonFieldRef } from '../../services/SharePointDataService';
import { getInitials } from './PeoplePicker';
import { ExpandableCellText } from './ExpandableCellText';

export interface IPersonDisplayProps {
  person: IPersonFieldRef | undefined;
  fallbackText: string | undefined;
}

// Mostra un Chip con Avatar quando il campo persona e' valorizzato (utente
// reale del tenant), altrimenti il testo del campo esistente (fallback per
// soggetti esterni al tenant o dati inseriti prima di questo collegamento,
// che puo' essere lungo: click per vederlo per intero se troncato).
export function PersonDisplay({ person, fallbackText }: IPersonDisplayProps): React.ReactElement {
  if (person && person.Title) {
    return <Chip size="small" avatar={<Avatar>{getInitials(person.Title)}</Avatar>} label={person.Title} />;
  }
  return <ExpandableCellText text={fallbackText} />;
}
