import { useMemo, useState } from 'react';
import type { ReactElement } from 'react';
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Tab,
  TabList,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { AddRegular, DeleteRegular, EditRegular, FolderRegular } from '@fluentui/react-icons';
import type {
  EnvironmentGroup,
  GovernanceRuleSet,
  PolicyRuleAssignment,
  RuleBasedPolicy,
} from '../types/admin.ts';
import type { Resource } from '../types/inventory.ts';
import ConfirmDialog from './ConfirmDialog.tsx';
import EmptyState from './EmptyState.tsx';
import { OperationsSkeleton, PageHeader } from './ui.tsx';
import { useMutation } from '../hooks/useMutation.tsx';
import { useOwners } from '../services/ownerCache.ts';
import {
  assignPolicyToGroup,
  createEnvironmentGroup,
  createGovernanceRuleSet,
  createRuleBasedPolicy,
  deleteEnvironmentGroup,
  deleteGovernanceRuleSet,
  updateEnvironmentGroup,
  updateGovernanceRuleSet,
  updateRuleBasedPolicy,
} from '../services/governanceMutations.ts';
import { addEnvironmentToGroup, removeEnvironmentFromGroup } from '../services/adminApi.ts';

interface EnvironmentGroupsViewProps {
  environments: Resource[];
  envGroups: EnvironmentGroup[];
  ruleBasedPolicies: RuleBasedPolicy[];
  ruleAssignments: PolicyRuleAssignment[];
  ruleSets: GovernanceRuleSet[];
  isLoading: boolean;
  error: string | null;
  onRefreshAdmin?: () => Promise<void>;
}

type EnvironmentGroupsTab = 'environmentGroups' | 'policies' | 'ruleSets';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '18px',
    padding: '28px 32px 32px',
    height: '100%',
    overflow: 'hidden',
    '@media (max-width: 768px)': {
      padding: tokens.spacingHorizontalM,
    },
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    minHeight: 0,
    flex: 1,
  },
  toolbar: {
    display: 'flex',
    justifyContent: 'flex-end',
  },
  tableWrapper: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },
  dialogSurface: {
    color: '#F4FBFD',
    border: '1px solid #29404F',
    borderRadius: 0,
    backgroundColor: '#0C141D',
    boxShadow: '0 24px 64px rgba(0, 0, 0, 0.62)',
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    minWidth: '600px',
  },
  th: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    textAlign: 'left',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground3,
    whiteSpace: 'nowrap',
  },
  td: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle',
  },
  monoCell: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
  },
  tabular: {
    fontVariantNumeric: 'tabular-nums',
  },
  centered: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100%',
  },
});

interface GroupFormDialogProps {
  open: boolean;
  initial?: EnvironmentGroup;
  isLoading: boolean;
  environments?: Resource[];
  memberIds?: Set<string>;
  isAddingEnv?: boolean;
  isRemovingEnv?: boolean;
  onAddEnv?: (envId: string) => void;
  onRemoveEnv?: (envId: string) => void;
  onSubmit: (displayName: string, description: string) => void;
  onCancel: () => void;
}

function GroupFormDialog({ open, initial, isLoading, environments, memberIds, isAddingEnv, isRemovingEnv, onAddEnv, onRemoveEnv, onSubmit, onCancel }: GroupFormDialogProps): ReactElement {
  const styles = useStyles();
  const [displayName, setDisplayName] = useState(initial?.displayName ?? '');
  const [description, setDescription] = useState(initial?.description ?? '');
  const [addingEnvId, setAddingEnvId] = useState('');
  const isEdit = Boolean(initial);

  useMemo(() => {
    setDisplayName(initial?.displayName ?? '');
    setDescription(initial?.description ?? '');
    setAddingEnvId('');
  }, [initial, open]);

  const members = useMemo(
    () => (environments ?? []).filter((e) => memberIds?.has(e.name)),
    [environments, memberIds],
  );
  const nonMembers = useMemo(
    () => (environments ?? []).filter((e) => !memberIds?.has(e.name)),
    [environments, memberIds],
  );
  const selectedEnv = nonMembers.find((e) => e.name === addingEnvId);

  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.dialogSurface} style={{ maxWidth: 'min(580px, 95vw)', width: 'min(580px, 95vw)' }}>
        <DialogBody>
          <DialogTitle>{isEdit ? 'Edit Environment Group' : 'New Environment Group'}</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
            <Field label="Display name" required>
              <Input value={displayName} onChange={(_, d) => setDisplayName(d.value)} placeholder="My Group" />
            </Field>
            <Field label="Description">
              <Textarea value={description} onChange={(_, d) => setDescription(d.value)} rows={3} />
            </Field>

            {isEdit && environments && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS }}>
                <Text weight="semibold" size={200} style={{ textTransform: 'uppercase', letterSpacing: '0.04em', color: tokens.colorNeutralForeground3 }}>
                  Environments ({members.length})
                </Text>

                {members.length === 0 ? (
                  <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>No environments in this group.</Text>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, maxHeight: '180px', overflowY: 'auto', border: '1px solid #29404F', borderRadius: 0, padding: tokens.spacingHorizontalS, backgroundColor: '#09121A' }}>
                    {members.map((env) => (
                      <div key={env.name} style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                        <Text style={{ flex: 1, fontSize: tokens.fontSizeBase200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {env.properties.displayName ?? env.name}
                        </Text>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>
                          {(env.properties as { environmentType?: string }).environmentType ?? ''}
                        </Text>
                        <Button
                          appearance="subtle" size="small" icon={<DeleteRegular />}
                          title="Remove from group"
                          disabled={isRemovingEnv || isAddingEnv}
                          onClick={() => onRemoveEnv?.(env.name)}
                        />
                      </div>
                    ))}
                  </div>
                )}

                {nonMembers.length > 0 && (
                  <div style={{ display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center' }}>
                    <Dropdown
                      placeholder="Add an environment…"
                      value={selectedEnv?.properties.displayName ?? selectedEnv?.name ?? ''}
                      selectedOptions={addingEnvId ? [addingEnvId] : []}
                      onOptionSelect={(_, d) => setAddingEnvId(d.optionValue ?? '')}
                      style={{ flex: 1 }}
                    >
                      {nonMembers.map((e) => (
                        <Option key={e.name} value={e.name} text={e.properties.displayName ?? e.name}>
                          <div>
                            <Text size={200} weight="semibold">{e.properties.displayName ?? e.name}</Text>
                            <Text size={200} style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
                              {(e.properties as { environmentType?: string }).environmentType ?? ''}
                            </Text>
                          </div>
                        </Option>
                      ))}
                    </Dropdown>
                    <Button
                      appearance="secondary" size="small" icon={<AddRegular />}
                      disabled={!addingEnvId || isAddingEnv || isRemovingEnv}
                      onClick={() => { if (addingEnvId) { onAddEnv?.(addingEnvId); setAddingEnvId(''); } }}
                    >
                      {isAddingEnv ? 'Adding…' : 'Add'}
                    </Button>
                  </div>
                )}
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading || !displayName.trim()} onClick={() => onSubmit(displayName.trim(), description.trim())}>
              {isLoading ? 'Saving…' : isEdit ? 'Save' : 'Create'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

function formatDate(value?: string): string {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

interface PolicyFormDialogProps {
  open: boolean;
  isLoading: boolean;
  onSubmit: (name: string) => void;
  onCancel: () => void;
}

function PolicyFormDialog({ open, isLoading, onSubmit, onCancel }: PolicyFormDialogProps): ReactElement {
  const styles = useStyles();
  const [name, setName] = useState('');
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.dialogSurface}>
        <DialogBody>
          <DialogTitle>New Rule-Based Policy</DialogTitle>
          <DialogContent>
            <Field label="Policy name" required>
              <Input value={name} onChange={(_, d) => setName(d.value)} placeholder="My Governance Policy" />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading || !name.trim()} onClick={() => { onSubmit(name.trim()); setName(''); }}>
              {isLoading ? 'Creating…' : 'Create'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

interface AssignPolicyDialogProps {
  open: boolean;
  policy: RuleBasedPolicy | null;
  envGroups: EnvironmentGroup[];
  assignedGroupIds: Set<string>;
  isLoading: boolean;
  onSubmit: (policyId: string, groupId: string) => void;
  onCancel: () => void;
}

function AssignPolicyDialog({ open, policy, envGroups, assignedGroupIds, isLoading, onSubmit, onCancel }: AssignPolicyDialogProps): ReactElement {
  const styles = useStyles();
  const [selectedGroupId, setSelectedGroupId] = useState('');
  const availableGroups = envGroups.filter((g) => !assignedGroupIds.has(g.id));
  const selectedGroup = envGroups.find((g) => g.id === selectedGroupId);
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) { setSelectedGroupId(''); onCancel(); } }}>
      <DialogSurface className={styles.dialogSurface}>
        <DialogBody>
          <DialogTitle>Assign Policy to Environment Group</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
            <Text>Policy: <strong>{policy?.name}</strong></Text>
            <Field label="Environment Group" required>
              <Dropdown
                placeholder={availableGroups.length === 0 ? 'All groups already assigned' : 'Select a group…'}
                value={selectedGroup?.displayName ?? ''}
                selectedOptions={selectedGroupId ? [selectedGroupId] : []}
                onOptionSelect={(_, data) => setSelectedGroupId(data.optionValue ?? '')}
                disabled={availableGroups.length === 0}
              >
                {availableGroups.map((g) => (
                  <Option key={g.id} value={g.id}>{g.displayName}</Option>
                ))}
              </Dropdown>
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={() => { setSelectedGroupId(''); onCancel(); }}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading || !selectedGroupId} onClick={() => { if (policy && selectedGroupId) { onSubmit(policy.id, selectedGroupId); setSelectedGroupId(''); } }}>
              {isLoading ? 'Assigning…' : 'Assign'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

interface PolicyEditDialogProps {
  open: boolean;
  policy: RuleBasedPolicy | null;
  availableRuleSets: GovernanceRuleSet[];
  isLoading: boolean;
  isSavingRuleSet: boolean;
  onSubmit: (policyId: string, name: string, ruleSets: RuleBasedPolicy['ruleSets']) => void;
  onSaveRuleSet: (rsId: string, data: Omit<GovernanceRuleSet, 'id' | 'lastModified'>) => void;
  onCancel: () => void;
}

function ruleSetToJson(rs: GovernanceRuleSet): string {
  return JSON.stringify({
    environmentFilter: { type: rs.environmentFilterType, values: rs.environmentFilterValues },
    parameters: rs.parameters.map((p) => ({ type: p.type, resourceType: p.resourceType, value: p.rules })),
  }, null, 2);
}

function emptyRuleSetJson(): string {
  return JSON.stringify({ environmentFilter: { type: '', values: [] }, parameters: [] }, null, 2);
}

function PolicyEditDialog({ open, policy, availableRuleSets, isLoading, isSavingRuleSet, onSubmit, onSaveRuleSet, onCancel }: PolicyEditDialogProps): ReactElement {
  const styles = useStyles();
  const [name, setName] = useState(policy?.name ?? '');
  const [currentRuleSets, setCurrentRuleSets] = useState(policy?.ruleSets ?? []);
  const [addingRuleSetId, setAddingRuleSetId] = useState('');
  const [expandedItems, setExpandedItems] = useState<string[]>([]);
  const [ruleSetEdits, setRuleSetEdits] = useState<Record<string, { json: string; error: string }>>({});
  const [savingRuleSetId, setSavingRuleSetId] = useState<string | null>(null);

  const ruleSetMap = useMemo(() => new Map(availableRuleSets.map((rs) => [rs.id, rs])), [availableRuleSets]);

  const policyId = policy?.id;
  useMemo(() => {
    setName(policy?.name ?? '');
    setCurrentRuleSets(policy?.ruleSets ?? []);
    setAddingRuleSetId('');
    setExpandedItems([]);
    setRuleSetEdits({});
    setSavingRuleSetId(null);
  }, [policy]);

  const assignedIds = new Set(currentRuleSets.map((rs) => rs.id));
  const unassigned = availableRuleSets.filter((rs) => rs.id && !assignedIds.has(rs.id));

  function removeRuleSet(id: string) {
    setCurrentRuleSets((prev) => prev.filter((rs) => rs.id !== id));
    setExpandedItems((prev) => prev.filter((x) => x !== id));
  }

  function addRuleSet() {
    if (!addingRuleSetId) return;
    setCurrentRuleSets((prev) => [...prev, { id: addingRuleSetId, version: '', inputs: {} }]);
    setAddingRuleSetId('');
  }

  function onAccordionToggle(_: unknown, data: { openItems: unknown[] }) {
    const newItems = data.openItems as string[];
    const opened = newItems.filter((id) => !expandedItems.includes(id));
    if (opened.length > 0) {
      setRuleSetEdits((prev) => {
        const next = { ...prev };
        opened.forEach((id) => {
          if (!next[id]) {
            const full = ruleSetMap.get(id);
            next[id] = { json: full ? ruleSetToJson(full) : emptyRuleSetJson(), error: '' };
          }
        });
        return next;
      });
    }
    setExpandedItems(newItems);
  }

  function updateJson(id: string, value: string) {
    setRuleSetEdits((prev) => ({ ...prev, [id]: { json: value, error: '' } }));
  }

  function handleSaveRuleSet(id: string) {
    const edit = ruleSetEdits[id];
    if (!edit) return;
    type ParsedRuleSet = { environmentFilter?: { type?: string; values?: { id: string; type: string; name: string }[] }; parameters?: { type: string; resourceType: string; value?: { id: string; value: string }[] }[] };
    let parsed: ParsedRuleSet;
    try {
      parsed = JSON.parse(edit.json) as ParsedRuleSet;
    } catch {
      setRuleSetEdits((prev) => ({ ...prev, [id]: { ...prev[id], error: 'Invalid JSON — please fix before saving.' } }));
      return;
    }
    const data: Omit<GovernanceRuleSet, 'id' | 'lastModified'> = {
      environmentFilterType: parsed.environmentFilter?.type ?? '',
      environmentFilterValues: parsed.environmentFilter?.values ?? [],
      parameters: (parsed.parameters ?? []).map((p) => ({ type: p.type, resourceType: p.resourceType, rules: p.value ?? [] })),
    };
    setSavingRuleSetId(id);
    onSaveRuleSet(id, data);
  }

  useMemo(() => { if (!isSavingRuleSet) setSavingRuleSetId(null); }, [isSavingRuleSet]);

  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.dialogSurface} style={{ maxWidth: 'min(640px, 95vw)', width: 'min(640px, 95vw)' }}>
        <DialogBody>
          <DialogTitle>Edit Policy</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
            <Field label="Policy name" required>
              <Input value={name} onChange={(_, d) => setName(d.value)} />
            </Field>
            <Field label="Rule Sets">
              <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }}>
                {currentRuleSets.length === 0 && (
                  <Text style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>No rule sets assigned.</Text>
                )}
                <Accordion
                  collapsible
                  multiple
                  openItems={expandedItems}
                  onToggle={onAccordionToggle}
                >
                  {currentRuleSets.map((rs) => {
                    const edit = ruleSetEdits[rs.id];
                    const isSaving = savingRuleSetId === rs.id && isSavingRuleSet;
                    return (
                      <AccordionItem key={rs.id} value={rs.id} style={{ border: '1px solid #29404F', borderRadius: 0, marginBottom: tokens.spacingVerticalXS, backgroundColor: '#09121A' }}>
                        <AccordionHeader
                          expandIconPosition="end"
                          style={{ paddingRight: tokens.spacingHorizontalXS }}
                          icon={null}
                        >
                          <span style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, flex: 1, minWidth: 0 }}>
                            <Text style={{ fontSize: tokens.fontSizeBase200, fontFamily: 'monospace', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 }}>{rs.id}</Text>
                            {rs.version && <Badge appearance="outline" size="small">{rs.version}</Badge>}
                            <Button
                              appearance="subtle" size="small" icon={<DeleteRegular />}
                              title="Remove rule set from policy"
                              onClick={(e) => { e.stopPropagation(); removeRuleSet(rs.id); }}
                            />
                          </span>
                        </AccordionHeader>
                        <AccordionPanel>
                          <div style={{ padding: `0 ${tokens.spacingHorizontalM} ${tokens.spacingVerticalM}`, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS }}>
                            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                              {ruleSetMap.has(rs.id) ? 'Edit rule set definition:' : 'Rule set not in tenant list — define manually:'}
                            </Text>
                            <Textarea
                              value={edit?.json ?? ''}
                              onChange={(_, d) => updateJson(rs.id, d.value)}
                              rows={10}
                              style={{ fontFamily: 'monospace', fontSize: tokens.fontSizeBase200 }}
                            />
                            {edit?.error && (
                              <Text style={{ color: tokens.colorPaletteRedForeground1, fontSize: tokens.fontSizeBase200 }}>{edit.error}</Text>
                            )}
                            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: tokens.spacingHorizontalS }}>
                              <Button
                                appearance="subtle" size="small"
                                onClick={() => {
                                  const full = ruleSetMap.get(rs.id);
                                  setRuleSetEdits((prev) => ({ ...prev, [rs.id]: { json: full ? ruleSetToJson(full) : emptyRuleSetJson(), error: '' } }));
                                }}
                              >Reset</Button>
                              <Button
                                appearance="primary" size="small"
                                disabled={isSaving}
                                onClick={() => handleSaveRuleSet(rs.id)}
                              >
                                {isSaving ? 'Saving…' : 'Save rule set'}
                              </Button>
                            </div>
                          </div>
                        </AccordionPanel>
                      </AccordionItem>
                    );
                  })}
                </Accordion>
                {unassigned.length > 0 && (
                  <div style={{ display: 'flex', gap: tokens.spacingHorizontalS, marginTop: tokens.spacingVerticalXS }}>
                    <Dropdown
                      placeholder="Add a rule set…"
                      value={addingRuleSetId || undefined}
                      selectedOptions={addingRuleSetId ? [addingRuleSetId] : []}
                      onOptionSelect={(_, d) => setAddingRuleSetId(d.optionValue ?? '')}
                      style={{ flex: 1 }}
                    >
                      {unassigned.map((rs) => <Option key={rs.id} value={rs.id}>{rs.id}</Option>)}
                    </Dropdown>
                    <Button appearance="secondary" size="small" icon={<AddRegular />} disabled={!addingRuleSetId} onClick={addRuleSet}>Add</Button>
                  </div>
                )}
              </div>
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading || !name.trim()} onClick={() => { if (policyId) onSubmit(policyId, name.trim(), currentRuleSets); }}>
              {isLoading ? 'Saving…' : 'Save'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

interface RuleSetFormDialogProps {
  open: boolean;
  initial: GovernanceRuleSet | null;
  template?: Omit<GovernanceRuleSet, 'id' | 'lastModified'> | null;
  envGroups: EnvironmentGroup[];
  isLoading: boolean;
  onSubmit: (data: Omit<GovernanceRuleSet, 'id' | 'lastModified'>, groupId?: string) => void;
  onCancel: () => void;
}

function RuleSetFormDialog({ open, initial, template, envGroups, isLoading, onSubmit, onCancel }: RuleSetFormDialogProps): ReactElement {
  const styles = useStyles();
  const isEdit = Boolean(initial);
  const [groupId, setGroupId] = useState('');
  const [jsonValue, setJsonValue] = useState('');
  const [jsonError, setJsonError] = useState('');

  useMemo(() => {
    const source = initial ?? template;
    if (source) {
      setJsonValue(JSON.stringify({
        environmentFilter: { type: source.environmentFilterType, values: source.environmentFilterValues },
        parameters: source.parameters.map((p) => ({ type: p.type, resourceType: p.resourceType, value: p.rules })),
      }, null, 2));
    } else {
      setJsonValue(JSON.stringify({ environmentFilter: { type: '', values: [] }, parameters: [] }, null, 2));
    }
    setGroupId('');
    setJsonError('');
  }, [initial, template, open]);

  function handleSave() {
    let parsed: { environmentFilter?: { type?: string; values?: { id: string; type: string; name: string }[] }; parameters?: { type: string; resourceType: string; value?: { id: string; value: string }[] }[] };
    try {
      parsed = JSON.parse(jsonValue) as typeof parsed;
    } catch {
      setJsonError('Invalid JSON — please fix before saving.');
      return;
    }
    const data: Omit<GovernanceRuleSet, 'id' | 'lastModified'> = {
      environmentFilterType: parsed.environmentFilter?.type ?? '',
      environmentFilterValues: parsed.environmentFilter?.values ?? [],
      parameters: (parsed.parameters ?? []).map((p) => ({ type: p.type, resourceType: p.resourceType, rules: p.value ?? [] })),
    };
    onSubmit(data, isEdit ? undefined : groupId);
  }

  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.dialogSurface} style={{ maxWidth: 'min(640px, 95vw)', width: 'min(640px, 95vw)' }}>
        <DialogBody>
          <DialogTitle>{isEdit ? 'Edit Rule Set' : template ? 'New Rule Set (from template)' : 'New Rule Set'}</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
            {!isEdit && (
              <Field label="Environment Group" required>
                <Dropdown
                  placeholder="Select a group…"
                  value={envGroups.find((g) => g.id === groupId)?.displayName ?? ''}
                  selectedOptions={groupId ? [groupId] : []}
                  onOptionSelect={(_, d) => setGroupId(d.optionValue ?? '')}
                >
                  {envGroups.map((g) => <Option key={g.id} value={g.id}>{g.displayName}</Option>)}
                </Dropdown>
              </Field>
            )}
            <Field label="Rule Set definition (JSON)" validationMessage={jsonError} validationState={jsonError ? 'error' : 'none'}>
              <Textarea
                value={jsonValue}
                onChange={(_, d) => { setJsonValue(d.value); setJsonError(''); }}
                rows={14}
                style={{ fontFamily: 'monospace', fontSize: tokens.fontSizeBase200 }}
              />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={isLoading} onClick={onCancel}>Cancel</Button>
            <Button appearance="primary" disabled={isLoading || (!isEdit && !groupId)} onClick={handleSave}>
              {isLoading ? 'Saving…' : isEdit ? 'Save' : 'Create'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

interface ExtractRuleSetDialogProps {
  open: boolean;
  policy: RuleBasedPolicy | null;
  ruleSets: GovernanceRuleSet[];
  onSelect: (template: Omit<GovernanceRuleSet, 'id' | 'lastModified'>) => void;
  onCancel: () => void;
}

function ExtractRuleSetDialog({ open, policy, ruleSets, onSelect, onCancel }: ExtractRuleSetDialogProps): ReactElement {
  const styles = useStyles();
  const ruleSetMap = useMemo(() => new Map(ruleSets.map((rs) => [rs.id, rs])), [ruleSets]);
  const policyRuleSets = policy?.ruleSets ?? [];

  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface className={styles.dialogSurface} style={{ maxWidth: 'min(600px, 95vw)' }}>
        <DialogBody>
          <DialogTitle>Extract rule as template from "{policy?.name}"</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }}>
            {policyRuleSets.length === 0 ? (
              <Text style={{ color: tokens.colorNeutralForeground3 }}>This policy has no rule sets to extract.</Text>
            ) : policyRuleSets.map((prs) => {
              const full = ruleSetMap.get(prs.id);
              const inputKeys = Object.keys(prs.inputs ?? {});

              return (
                <div key={prs.id} style={{ border: '1px solid #29404F', borderRadius: 0, overflow: 'hidden', backgroundColor: '#09121A' }}>
                  <div style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, background: tokens.colorNeutralBackground3, display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                    <Text style={{ fontFamily: 'monospace', fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold }}>{prs.id || '(no id)'}</Text>
                    {prs.version && <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>v{prs.version}</Text>}
                    {!full && <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, marginLeft: 'auto' }}>not in tenant rule set list</Text>}
                  </div>

                  {full && full.parameters.length > 0 && full.parameters.map((param, i) => (
                    <div key={`${prs.id}-p-${i}`} style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM, padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, borderTop: `1px solid ${tokens.colorNeutralStroke2}` }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase200 }}>{param.type || '(unnamed)'}</Text>
                        {param.resourceType && <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}> · {param.resourceType}</Text>}
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, display: 'block' }}>{param.rules.length} rule(s)</Text>
                      </div>
                      <Button appearance="primary" size="small" onClick={() =>
                        onSelect({ environmentFilterType: full.environmentFilterType, environmentFilterValues: full.environmentFilterValues, parameters: [param] })
                      }>
                        Use as template
                      </Button>
                    </div>
                  ))}

                  {!full && inputKeys.length > 0 && inputKeys.map((key) => (
                    <div key={`${prs.id}-i-${key}`} style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM, padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, borderTop: `1px solid ${tokens.colorNeutralStroke2}` }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <Text style={{ fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase200 }}>{key}</Text>
                        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, display: 'block' }}>{JSON.stringify(prs.inputs[key]).slice(0, 80)}</Text>
                      </div>
                      <Button appearance="primary" size="small" onClick={() =>
                        onSelect({ environmentFilterType: '', environmentFilterValues: [], parameters: [{ type: key, resourceType: '', rules: [] }] })
                      }>
                        Use as template
                      </Button>
                    </div>
                  ))}

                  {!full && inputKeys.length === 0 && (
                    <div style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, borderTop: `1px solid ${tokens.colorNeutralStroke2}` }}>
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>No rule details available for this rule set.</Text>
                    </div>
                  )}

                  {full && full.parameters.length === 0 && (
                    <div style={{ padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`, borderTop: `1px solid ${tokens.colorNeutralStroke2}` }}>
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>This rule set has no parameters to extract.</Text>
                    </div>
                  )}
                </div>
              );
            })}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onCancel}>Close</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

function renderEnvironmentGroupsTable(
  styles: ReturnType<typeof useStyles>,
  envGroups: EnvironmentGroup[],
  pendingGroupId: string | null,
  ownerNames: Map<string, string>,
  envsByGroup: Map<string, Resource[]>,
  onEdit: (group: EnvironmentGroup) => void,
  onDelete: (group: EnvironmentGroup) => void,
): ReactElement {
  if (envGroups.length === 0) {
    return (
      <EmptyState
        icon={<FolderRegular />}
        title="No environment groups found"
        subtitle="Create a group to organize environments and apply shared governance policies."
      />
    );
  }
  return (
    <div className={styles.tableWrapper}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th className={styles.th}>Display Name</th>
            <th className={styles.th}>Description</th>
            <th className={styles.th}>Created By</th>
            <th className={styles.th}>Created Time</th>
            <th className={styles.th}>Environments</th>
            <th className={styles.th} style={{ width: '80px' }}></th>
          </tr>
        </thead>
        <tbody>
          {envGroups.map((group) => {
              const envCount = envsByGroup.get(group.id)?.length ?? 0;
              return (
                <tr key={group.id}>
                  <td className={styles.td}>{group.displayName}</td>
                  <td className={styles.td}>{group.description || '—'}</td>
                  <td className={styles.td}>{ownerNames.get(group.createdBy.id?.toLowerCase()) ?? group.createdBy.displayName ?? group.createdBy.email ?? group.createdBy.id}</td>
                  <td className={styles.td}>{formatDate(group.createdTime)}</td>
                  <td className={styles.td}>
                    <Badge appearance="tint" color={envCount > 0 ? 'brand' : 'informative'} size="small">
                      {envCount}
                    </Badge>
                  </td>
                  <td className={styles.td}>
                    <span style={{ display: 'flex', gap: tokens.spacingHorizontalXS }}>
                      <Button
                        appearance="subtle" size="small"
                        icon={<EditRegular />}
                        disabled={pendingGroupId === group.id}
                        title="Edit"
                        onClick={() => onEdit(group)}
                      />
                      <Button
                        appearance="subtle" size="small"
                        icon={<DeleteRegular />}
                        disabled={pendingGroupId === group.id}
                        title="Delete"
                        style={{ color: tokens.colorStatusDangerForeground1 }}
                        onClick={() => onDelete(group)}
                      />
                    </span>
                  </td>
                </tr>
              );
            })}
        </tbody>
      </table>
    </div>
  );
}

export default function EnvironmentGroupsView({
  environments,
  envGroups,
  ruleBasedPolicies,
  ruleAssignments,
  ruleSets,
  isLoading,
  error,
  onRefreshAdmin,
}: EnvironmentGroupsViewProps): ReactElement {
  const styles = useStyles();
  const [activeTab, setActiveTab] = useState<EnvironmentGroupsTab>('environmentGroups');
  const [groupFormOpen, setGroupFormOpen] = useState(false);
  const [editingGroup, setEditingGroup] = useState<EnvironmentGroup | null>(null);
  const [deleteGroup, setDeleteGroup] = useState<EnvironmentGroup | null>(null);
  const [pendingGroupId, setPendingGroupId] = useState<string | null>(null);
  const [policyFormOpen, setPolicyFormOpen] = useState(false);
  const [assignPolicyTarget, setAssignPolicyTarget] = useState<RuleBasedPolicy | null>(null);
  const [editingPolicy, setEditingPolicy] = useState<RuleBasedPolicy | null>(null);
  const [ruleSetFormOpen, setRuleSetFormOpen] = useState(false);
  const [editingRuleSet, setEditingRuleSet] = useState<GovernanceRuleSet | null>(null);
  const [deleteRuleSetTarget, setDeleteRuleSetTarget] = useState<GovernanceRuleSet | null>(null);
  const [extractingFromPolicy, setExtractingFromPolicy] = useState<RuleBasedPolicy | null>(null);
  const [ruleSetTemplate, setRuleSetTemplate] = useState<Omit<GovernanceRuleSet, 'id' | 'lastModified'> | null>(null);

  const groupOwnerGuids = useMemo(
    () => envGroups.map((g) => g.createdBy?.id as string | undefined),
    [envGroups],
  );
  const groupOwnerNames = useOwners(groupOwnerGuids);

  const { execute: execCreate, isLoading: isCreating } = useMutation(createEnvironmentGroup, {
    successMessage: 'Environment group created.',
    onSuccess: () => { setGroupFormOpen(false); void onRefreshAdmin?.(); },
  });
  const { execute: execUpdate, isLoading: isUpdating } = useMutation(updateEnvironmentGroup, {
    successMessage: 'Environment group updated.',
    onSuccess: () => { setEditingGroup(null); void onRefreshAdmin?.(); },
  });
  const { execute: execDelete } = useMutation(deleteEnvironmentGroup, {
    successMessage: 'Environment group deleted.',
    onSuccess: () => { setPendingGroupId(null); void onRefreshAdmin?.(); },
  });
  const { execute: execCreatePolicy, isLoading: isCreatingPolicy } = useMutation(createRuleBasedPolicy, {
    successMessage: 'Policy created.',
    onSuccess: () => { setPolicyFormOpen(false); void onRefreshAdmin?.(); },
  });
  const { execute: execAssignPolicy, isLoading: isAssigningPolicy } = useMutation(assignPolicyToGroup, {
    successMessage: 'Policy assigned to environment group.',
    onSuccess: () => { setAssignPolicyTarget(null); void onRefreshAdmin?.(); },
  });
  const { execute: execUpdatePolicy, isLoading: isUpdatingPolicy } = useMutation(updateRuleBasedPolicy, {
    successMessage: 'Policy updated.',
    onSuccess: () => { setEditingPolicy(null); void onRefreshAdmin?.(); },
  });
  const { execute: execCreateRuleSet, isLoading: isCreatingRuleSet } = useMutation(createGovernanceRuleSet, {
    successMessage: 'Rule set created.',
    onSuccess: () => { setRuleSetFormOpen(false); setEditingRuleSet(null); setRuleSetTemplate(null); void onRefreshAdmin?.(); },
  });
  const { execute: execUpdateRuleSet, isLoading: isUpdatingRuleSet } = useMutation(updateGovernanceRuleSet, {
    successMessage: 'Rule set updated.',
    onSuccess: () => { setEditingRuleSet(null); void onRefreshAdmin?.(); },
  });
  const { execute: execDeleteRuleSet } = useMutation(deleteGovernanceRuleSet, {
    successMessage: 'Rule set deleted.',
    onSuccess: () => { setDeleteRuleSetTarget(null); void onRefreshAdmin?.(); },
  });
  const { execute: execAddEnvToGroup, isLoading: isAddingEnvToGroup } = useMutation(
    ({ groupId, envId }: { groupId: string; envId: string }) => addEnvironmentToGroup(groupId, envId),
    {
      successMessage: 'Environment added to group.',
      onSuccess: () => void onRefreshAdmin?.(),
    },
  );
  const { execute: execRemoveEnvFromGroup, isLoading: isRemovingEnvFromGroup } = useMutation(
    ({ groupId, envId }: { groupId: string; envId: string }) => removeEnvironmentFromGroup(groupId, envId),
    {
      successMessage: 'Environment removed from group.',
      onSuccess: () => void onRefreshAdmin?.(),
    },
  );

  const groupMap = useMemo(() => new Map(envGroups.map((g) => [g.id, g.displayName])), [envGroups]);
  const envsByGroup = useMemo(
    () => {
      const map = new Map<string, Resource[]>();
      for (const env of environments) {
        const gId = env.properties['environmentGroupId'] as string | undefined;
        if (gId) {
          const arr = map.get(gId) ?? [];
          arr.push(env);
          map.set(gId, arr);
        }
      }
      return map;
    },
    [environments, envGroups],
  );

  const content = useMemo(() => {
    if (isLoading) {
      return (
        <div className={styles.centered}>
          <OperationsSkeleton />
        </div>
      );
    }

    if (error) {
      return (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      );
    }

    switch (activeTab) {
      case 'policies': {
        const assignmentsByPolicy = new Map<string, PolicyRuleAssignment[]>();
        for (const a of ruleAssignments) {
          if (!a.policyId) continue;
          const list = assignmentsByPolicy.get(a.policyId) ?? [];
          list.push(a);
          assignmentsByPolicy.set(a.policyId, list);
        }
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, flex: 1, minHeight: 0 }}>
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <Button appearance="primary" size="small" icon={<AddRegular />} onClick={() => setPolicyFormOpen(true)}>
                New Policy
              </Button>
            </div>
            <div className={styles.tableWrapper}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th className={styles.th}>Name</th>
                    <th className={styles.th}>Rule Sets</th>
                    <th className={styles.th}>Assigned To</th>
                    <th className={styles.th}>Last Modified</th>
                    <th className={styles.th} style={{ width: '80px' }}></th>
                  </tr>
                </thead>
                <tbody>
                  {ruleBasedPolicies.length === 0 ? (
                    <tr>
                      <td className={styles.td} colSpan={5} style={{ textAlign: 'center' }}>No rule-based policies found.</td>
                    </tr>
                  ) : ruleBasedPolicies.map((policy) => {
                    const assignments = assignmentsByPolicy.get(policy.id) ?? [];
                    const groupAssignments = assignments.filter((a) => a.resourceType === 'EnvironmentGroup');
                    return (
                      <tr key={policy.id}>
                        <td className={styles.td}>{policy.name || '—'}</td>
                        <td className={styles.td}>{policy.ruleSetCount}</td>
                        <td className={styles.td}>
                          <span style={{ display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS }}>
                            {groupAssignments.length === 0
                              ? <span style={{ color: tokens.colorNeutralForeground3 }}>—</span>
                              : groupAssignments.map((a) => (
                                <Badge key={a.resourceId} appearance="filled" color="informative">
                                  {groupMap.get(a.resourceId) ?? a.resourceId}
                                </Badge>
                              ))
                            }
                          </span>
                        </td>
                        <td className={styles.td}>{formatDate(policy.lastModified)}</td>
                        <td className={styles.td}>
                          <span style={{ display: 'flex', gap: tokens.spacingHorizontalXS }}>
                            <Button appearance="subtle" size="small" icon={<EditRegular />} title="Edit policy" onClick={() => setEditingPolicy(policy)} />
                            <Button appearance="subtle" size="small" icon={<AddRegular />} title="Assign to Environment Group" onClick={() => setAssignPolicyTarget(policy)} />
                            {policy.ruleSets.length > 0 && (
                              <Button appearance="subtle" size="small" title="Extract rule set as template" onClick={() => setExtractingFromPolicy(policy)}>Extract</Button>
                            )}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        );
      }
      case 'ruleSets':
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, flex: 1, minHeight: 0 }}>
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <Button appearance="primary" size="small" icon={<AddRegular />} onClick={() => { setEditingRuleSet(null); setRuleSetFormOpen(true); }}>
                New Rule Set
              </Button>
            </div>
            <div className={styles.tableWrapper}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th className={styles.th}>ID</th>
                    <th className={styles.th}>Filter Type</th>
                    <th className={styles.th}>Environments</th>
                    <th className={styles.th}>Parameters</th>
                    <th className={styles.th}>Last Modified</th>
                    <th className={styles.th} style={{ width: '80px' }}></th>
                  </tr>
                </thead>
                <tbody>
                  {ruleSets.length === 0 ? (
                    <tr><td className={styles.td} colSpan={6} style={{ textAlign: 'center' }}>No rule sets found.</td></tr>
                  ) : ruleSets.map((rs) => (
                    <tr key={rs.id}>
                      <td className={styles.td} style={{ fontFamily: 'monospace', fontSize: tokens.fontSizeBase200 }}>{rs.id}</td>
                      <td className={styles.td}>{rs.environmentFilterType || '—'}</td>
                      <td className={styles.td}>{rs.environmentFilterValues.length}</td>
                      <td className={styles.td}>{rs.parameters.length}</td>
                      <td className={styles.td}>{formatDate(rs.lastModified)}</td>
                      <td className={styles.td}>
                        <span style={{ display: 'flex', gap: tokens.spacingHorizontalXS }}>
                          <Button appearance="subtle" size="small" icon={<EditRegular />} title="Edit" onClick={() => { setEditingRuleSet(rs); setRuleSetFormOpen(true); }} />
                          <Button appearance="subtle" size="small" icon={<DeleteRegular />} title="Delete" onClick={() => setDeleteRuleSetTarget(rs)} />
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      case 'environmentGroups':
      default:
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, flex: 1, minHeight: 0 }}>
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <Button appearance="primary" size="small" icon={<AddRegular />} onClick={() => setGroupFormOpen(true)}>
                New Group
              </Button>
            </div>
            {renderEnvironmentGroupsTable(styles, envGroups, pendingGroupId, groupOwnerNames, envsByGroup, (g) => setEditingGroup(g), (g) => setDeleteGroup(g))}
          </div>
        );
    }
  }, [activeTab, envGroups, envsByGroup, error, groupMap, groupOwnerNames, isLoading, pendingGroupId, ruleAssignments, ruleBasedPolicies, ruleSets, styles]);

  return (
    <div className={styles.root}>
      <PageHeader
        title="Environment groups"
        description="Organize environments and apply reusable governance policies and rule sets."
      />

      <div className={styles.body}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, data) => setActiveTab(data.value as EnvironmentGroupsTab)}
        >
          <Tab value="environmentGroups">Environment Groups</Tab>
          <Tab value="policies">Rule-Based Policies</Tab>
          <Tab value="ruleSets">Rule Sets</Tab>
        </TabList>

        {content}
      </div>

      <GroupFormDialog
        open={groupFormOpen}
        isLoading={isCreating}
        onSubmit={(name, desc) => void execCreate(name, desc)}
        onCancel={() => setGroupFormOpen(false)}
      />

      {editingGroup && (
        <GroupFormDialog
          open
          initial={editingGroup}
          isLoading={isUpdating}
          environments={environments}
          memberIds={new Set((envsByGroup.get(editingGroup.id) ?? []).map((e) => e.name))}
          isAddingEnv={isAddingEnvToGroup}
          isRemovingEnv={isRemovingEnvFromGroup}
          onAddEnv={(envId) => void execAddEnvToGroup({ groupId: editingGroup.id, envId })}
          onRemoveEnv={(envId) => void execRemoveEnvFromGroup({ groupId: editingGroup.id, envId })}
          onSubmit={(name, desc) => void execUpdate(editingGroup.id, name, desc)}
          onCancel={() => setEditingGroup(null)}
        />
      )}

      <ConfirmDialog
        open={Boolean(deleteGroup)}
        title="Delete Environment Group"
        message={`Delete group "${deleteGroup?.displayName}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        onConfirm={() => {
          if (deleteGroup) {
            setPendingGroupId(deleteGroup.id);
            void execDelete(deleteGroup.id);
          }
          setDeleteGroup(null);
        }}
        onCancel={() => setDeleteGroup(null)}
      />

      <PolicyFormDialog
        open={policyFormOpen}
        isLoading={isCreatingPolicy}
        onSubmit={(name) => void execCreatePolicy(name)}
        onCancel={() => setPolicyFormOpen(false)}
      />

      {assignPolicyTarget && (
        <AssignPolicyDialog
          open
          policy={assignPolicyTarget}
          envGroups={envGroups}
          assignedGroupIds={new Set(
            ruleAssignments
              .filter((a) => a.policyId === assignPolicyTarget.id && a.resourceType === 'EnvironmentGroup')
              .map((a) => a.resourceId),
          )}
          isLoading={isAssigningPolicy}
          onSubmit={(policyId, groupId) => void execAssignPolicy(policyId, groupId)}
          onCancel={() => setAssignPolicyTarget(null)}
        />
      )}

      {editingPolicy && (
        <PolicyEditDialog
          open
          policy={editingPolicy}
          availableRuleSets={ruleSets}
          isLoading={isUpdatingPolicy}
          isSavingRuleSet={isUpdatingRuleSet}
          onSubmit={(policyId, name, policyRuleSets) => void execUpdatePolicy(policyId, name, policyRuleSets)}
          onSaveRuleSet={(rsId, data) => void execUpdateRuleSet(rsId, data)}
          onCancel={() => setEditingPolicy(null)}
        />
      )}

      <RuleSetFormDialog
        open={ruleSetFormOpen}
        initial={editingRuleSet}
        template={ruleSetTemplate}
        envGroups={envGroups}
        isLoading={editingRuleSet ? isUpdatingRuleSet : isCreatingRuleSet}
        onSubmit={(data, groupId) => {
          if (editingRuleSet) void execUpdateRuleSet(editingRuleSet.id, data);
          else if (groupId) void execCreateRuleSet(groupId, data);
        }}
        onCancel={() => { setRuleSetFormOpen(false); setEditingRuleSet(null); setRuleSetTemplate(null); }}
      />

      <ExtractRuleSetDialog
        open={Boolean(extractingFromPolicy)}
        policy={extractingFromPolicy}
        ruleSets={ruleSets}
        onSelect={(template) => { setExtractingFromPolicy(null); setEditingRuleSet(null); setRuleSetTemplate(template); setRuleSetFormOpen(true); }}
        onCancel={() => setExtractingFromPolicy(null)}
      />

      <ConfirmDialog
        open={Boolean(deleteRuleSetTarget)}
        title="Delete Rule Set"
        message={`Delete rule set "${deleteRuleSetTarget?.id}"? This action cannot be undone.`}
        confirmLabel="Delete"
        isDangerous
        onConfirm={() => { if (deleteRuleSetTarget) void execDeleteRuleSet(deleteRuleSetTarget.id); setDeleteRuleSetTarget(null); }}
        onCancel={() => setDeleteRuleSetTarget(null)}
      />
    </div>
  );
}
