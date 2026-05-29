// ─── Flow Definition Types ────────────────────────────────────────────────────

export interface FlowDefinition {
  triggers?: Record<string, FlowTrigger>;
  actions?: Record<string, FlowAction>;
  description?: string;
  contentVersion?: string;
}

export interface FlowTrigger {
  type?: string;
  kind?: string;
  description?: string;
  recurrence?: {
    frequency?: string;
    interval?: number;
  };
  conditions?: Array<{ expression?: string }>;
  runtimeConfiguration?: {
    concurrency?: { runs?: number };
    paginationPolicy?: { minimumItemCount?: number };
  };
  operationOptions?: string;
  inputs?: Record<string, unknown>;
}

export interface FlowAction {
  type?: string;
  kind?: string;
  inputs?: Record<string, unknown>;
  runAfter?: Record<string, string[]>;
  runtimeConfiguration?: {
    concurrency?: { repetitions?: number };
    paginationPolicy?: { minimumItemCount?: number };
    secureInputs?: boolean;
    secureOutputs?: boolean;
  };
  description?: string;
  metadata?: Record<string, unknown>;
  // Composite action children
  actions?: Record<string, FlowAction>;
  else?: { actions?: Record<string, FlowAction> };
  cases?: Record<string, { actions?: Record<string, FlowAction> }>;
  default?: { actions?: Record<string, FlowAction> };
  foreach?: string;
  expression?: unknown;
  operationOptions?: string;
  limit?: { timeout?: string };
}

// ─── Analysis Result ──────────────────────────────────────────────────────────

export type AnalysisSeverity = 'critical' | 'warning' | 'info';

export interface AnalysisResult {
  id: string;
  severity: AnalysisSeverity;
  title: string;
  description: string;
  recommendation: string;
  affectedItems?: string[];
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Recursively collect all actions (flat list) from nested scopes, conditions, loops, etc. */
function collectAllActions(
  actions: Record<string, FlowAction> | undefined,
  out: Array<[string, FlowAction]> = [],
): Array<[string, FlowAction]> {
  if (!actions) return out;
  for (const [name, action] of Object.entries(actions)) {
    out.push([name, action]);
    collectAllActions(action.actions, out);
    collectAllActions(action.else?.actions, out);
    if (action.cases) {
      for (const c of Object.values(action.cases)) collectAllActions(c.actions, out);
    }
    collectAllActions(action.default?.actions, out);
  }
  return out;
}

/** Collect actions with their nesting depth. */
function collectActionsWithDepth(
  actions: Record<string, FlowAction> | undefined,
  depth = 0,
  out: Array<[string, FlowAction, number]> = [],
): Array<[string, FlowAction, number]> {
  if (!actions) return out;
  for (const [name, action] of Object.entries(actions)) {
    out.push([name, action, depth]);
    collectActionsWithDepth(action.actions, depth + 1, out);
    collectActionsWithDepth(action.else?.actions, depth + 1, out);
    if (action.cases) {
      for (const c of Object.values(action.cases)) collectActionsWithDepth(c.actions, depth + 1, out);
    }
    collectActionsWithDepth(action.default?.actions, depth + 1, out);
  }
  return out;
}

const DEFAULT_NAME_PATTERNS = [
  /^Apply_to_each(_\d+)?$/i,
  /^Apply_to_each_\d+$/i,
  /^Condition(_\d+)?$/i,
  /^Do_until(_\d+)?$/i,
  /^Switch(_\d+)?$/i,
  /^Scope(_\d+)?$/i,
  /^Compose(_\d+)?$/i,
  /^Parse_JSON(_\d+)?$/i,
  /^Initialize_variable(_\d+)?$/i,
  /^Set_variable(_\d+)?$/i,
  /^Append_to_array_variable(_\d+)?$/i,
  /^Increment_variable(_\d+)?$/i,
  /^Send_an_email(_\d+)?$/i,
  /^Send_an_email_\(V\d+\)(_\d+)?$/i,
  /^HTTP(_\d+)?$/i,
  /^Response(_\d+)?$/i,
  /^Terminate(_\d+)?$/i,
];

function isDefaultName(name: string): boolean {
  return DEFAULT_NAME_PATTERNS.some((p) => p.test(name));
}

const SECRET_KEYWORDS = ['password', 'secret', 'token', 'apikey', 'api_key', 'key', 'credential', 'auth', 'bearer'];

function inputsLookSensitive(inputs: Record<string, unknown> | undefined): boolean {
  if (!inputs) return false;
  const str = JSON.stringify(inputs).toLowerCase();
  return SECRET_KEYWORDS.some((kw) => str.includes(kw));
}

// ─── Individual Checks ────────────────────────────────────────────────────────

function checkErrorHandling(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  // Top-level scope or actions that run on failure indicate error handling is present
  const hasScopeWithErrorRunAfter = allActions.some(([, a]) => {
    if (a.type?.toLowerCase() !== 'scope') return false;
    // A scope that runs after another action's failure
    const ra = a.runAfter ?? {};
    return Object.values(ra).some((statuses) =>
      statuses.some((s) => ['Failed', 'TimedOut', 'Skipped'].includes(s))
    );
  });

  const hasTerminateOnFailure = allActions.some(([, a]) => {
    if (a.type?.toLowerCase() !== 'terminate') return false;
    const ra = a.runAfter ?? {};
    return Object.values(ra).some((statuses) =>
      statuses.some((s) => ['Failed', 'TimedOut'].includes(s))
    );
  });

  if (hasScopeWithErrorRunAfter || hasTerminateOnFailure) return null;

  // Check how many actions handle failure runAfter
  const actionsHandlingFailure = allActions.filter(([, a]) => {
    const ra = a.runAfter ?? {};
    return Object.values(ra).some((statuses) =>
      statuses.some((s) => ['Failed', 'TimedOut', 'Skipped'].includes(s))
    );
  });

  if (actionsHandlingFailure.length > 0) return null;

  return {
    id: 'no-error-handling',
    severity: 'warning',
    title: 'No error handling detected',
    description: 'This flow does not appear to handle action failures. If any step fails, the flow will fail silently without notification or cleanup.',
    recommendation: 'Add a "Scope" action around the main logic and a second scope that runs on failure (set runAfter to Failed/TimedOut) to send notifications or perform cleanup.',
  };
}

function checkTriggerConditions(triggers: Record<string, FlowTrigger>): AnalysisResult | null {
  const pollingTriggers = Object.entries(triggers).filter(([, t]) => {
    const type = (t.type ?? '').toLowerCase();
    const kind = (t.kind ?? '').toLowerCase();
    return (
      type === 'recurrence' ||
      type === 'apiconnection' ||
      (type === 'request' && kind === 'button') ||
      kind === 'polling'
    );
  });

  const withoutConditions = pollingTriggers.filter(([, t]) => !t.conditions || t.conditions.length === 0);
  if (withoutConditions.length === 0) return null;

  return {
    id: 'no-trigger-conditions',
    severity: 'info',
    title: 'No trigger conditions set',
    description: 'Polling or connector triggers without conditions run every time the trigger fires, even when no action is needed.',
    recommendation: 'Add trigger conditions to filter events at source and reduce unnecessary flow runs — especially for high-frequency triggers.',
    affectedItems: withoutConditions.map(([name]) => name),
  };
}

function checkHighFrequencyRecurrence(triggers: Record<string, FlowTrigger>): AnalysisResult | null {
  const highFreq = Object.entries(triggers).filter(([, t]) => {
    if (t.type?.toLowerCase() !== 'recurrence') return false;
    const freq = (t.recurrence?.frequency ?? '').toLowerCase();
    const interval = t.recurrence?.interval ?? 1;
    return (freq === 'minute' && interval < 5) || freq === 'second';
  });

  if (highFreq.length === 0) return null;

  return {
    id: 'high-frequency-recurrence',
    severity: 'warning',
    title: 'Very high-frequency recurrence trigger',
    description: 'This flow runs very frequently (every few seconds or minutes), which can consume significant API call quota and increase costs.',
    recommendation: 'Review whether this frequency is truly needed. Consider pushing changes via webhooks/instant triggers instead of polling.',
    affectedItems: highFreq.map(([name]) => name),
  };
}

function checkConcurrencyControl(triggers: Record<string, FlowTrigger>): AnalysisResult | null {
  const triggersWithoutLimit = Object.entries(triggers).filter(([, t]) => {
    const type = (t.type ?? '').toLowerCase();
    if (!['recurrence', 'apiconnection', 'request'].includes(type)) return false;
    const concurrency = t.runtimeConfiguration?.concurrency?.runs;
    return concurrency === undefined || concurrency === null;
  });

  if (triggersWithoutLimit.length === 0) return null;

  return {
    id: 'no-concurrency-limit',
    severity: 'info',
    title: 'Trigger concurrency limit not set',
    description: 'Without a concurrency limit, multiple instances of this flow can run simultaneously. This can cause race conditions when writing to shared data sources.',
    recommendation: 'Set a concurrency limit on the trigger (1 if sequential processing is required, or a reasonable limit like 5–10) to avoid data corruption.',
    affectedItems: triggersWithoutLimit.map(([name]) => name),
  };
}

function checkDefaultActionNames(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const defaults = allActions
    .filter(([name]) => isDefaultName(name))
    .map(([name]) => name);

  if (defaults.length < 3) return null;

  return {
    id: 'default-action-names',
    severity: 'info',
    title: 'Many actions have default names',
    description: `${defaults.length} actions still use auto-generated names, making the flow harder to read, maintain, and debug.`,
    recommendation: 'Rename actions with descriptive names that explain what each step does. This greatly improves maintainability.',
    affectedItems: defaults.slice(0, 10),
  };
}

function checkHttpWithoutTimeout(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const httpNoTimeout = allActions
    .filter(([, a]) => a.type?.toLowerCase() === 'http' && !a.limit?.timeout)
    .map(([name]) => name);

  if (httpNoTimeout.length === 0) return null;

  return {
    id: 'http-no-timeout',
    severity: 'warning',
    title: 'HTTP actions without timeout',
    description: `${httpNoTimeout.length} HTTP action(s) have no timeout configured. If the remote endpoint is slow or unreachable, these steps can hang for up to the default 2-minute limit.`,
    recommendation: 'Set an explicit timeout (e.g., PT30S) on HTTP actions to fail fast and trigger retry/error handling logic.',
    affectedItems: httpNoTimeout.slice(0, 10),
  };
}

function checkSensitiveDataExposure(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const exposed = allActions.filter(([, a]) => {
    if (!inputsLookSensitive(a.inputs)) return false;
    const rt = a.runtimeConfiguration;
    return !rt?.secureInputs;
  }).map(([name]) => name);

  if (exposed.length === 0) return null;

  return {
    id: 'sensitive-data-exposure',
    severity: 'warning',
    title: 'Possible sensitive data in unprotected inputs',
    description: `${exposed.length} action(s) appear to reference sensitive values (passwords, tokens, keys) without enabling "Secure Inputs".`,
    recommendation: 'Enable "Secure Inputs" (and "Secure Outputs") on actions that handle credentials or secrets to prevent them from appearing in run history.',
    affectedItems: exposed.slice(0, 10),
  };
}

function checkForEachConcurrency(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const sequential = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'foreach') return false;
    const repetitions = a.runtimeConfiguration?.concurrency?.repetitions;
    return repetitions === undefined || repetitions === null;
  }).map(([name]) => name);

  if (sequential.length === 0) return null;

  return {
    id: 'foreach-no-concurrency',
    severity: 'info',
    title: '"Apply to each" loops run sequentially',
    description: `${sequential.length} loop(s) run one item at a time. For large collections this can be very slow.`,
    recommendation: 'Enable "concurrency" on "Apply to each" actions (degree of parallelism 2–50) when the loop body does not depend on shared state between iterations.',
    affectedItems: sequential,
  };
}

function checkPaginationOnListActions(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  // Find connector list actions (apiconnection type) that don't use pagination
  const listLike = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'apiconnection') return false;
    const op = (a.inputs?.['operationId'] as string | undefined ?? '').toLowerCase();
    const path = (a.inputs?.['path'] as string | undefined ?? '').toLowerCase();
    return op.includes('list') || op.includes('get_items') || path.includes('items') || path.includes('rows');
  });

  const noPagination = listLike.filter(([, a]) => {
    const policy = a.runtimeConfiguration?.paginationPolicy;
    return !policy || !policy.minimumItemCount;
  }).map(([name]) => name);

  if (noPagination.length === 0) return null;

  return {
    id: 'no-pagination',
    severity: 'info',
    title: 'List actions may not retrieve all items',
    description: `${noPagination.length} "list" action(s) do not have pagination enabled. By default, most connectors return only the first page of results (e.g., 100 items).`,
    recommendation: 'Enable pagination in action settings if you expect more items than the default page size, otherwise you may silently process an incomplete dataset.',
    affectedItems: noPagination.slice(0, 10),
  };
}

function checkFlowDescription(definition: FlowDefinition): AnalysisResult | null {
  if (definition.description && definition.description.trim().length > 0) return null;
  return {
    id: 'no-flow-description',
    severity: 'info',
    title: 'Flow has no description',
    description: 'This flow does not have a description explaining its purpose, owner, or expected behavior.',
    recommendation: 'Add a clear description to the flow covering: what it does, who owns it, what triggers it, and any dependencies.',
  };
}

// ─── NEW CHECKS ───────────────────────────────────────────────────────────────

function checkNestedForEach(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  // ForEach actions whose children also contain a ForEach
  const nestedLoops = allActions.filter(([, outer]) => {
    if (outer.type?.toLowerCase() !== 'foreach') return false;
    const children = collectAllActions(outer.actions);
    return children.some(([, inner]) => inner.type?.toLowerCase() === 'foreach');
  }).map(([name]) => name);

  if (nestedLoops.length === 0) return null;

  return {
    id: 'nested-foreach',
    severity: 'warning',
    title: 'Nested "Apply to each" loops detected',
    description: `${nestedLoops.length} loop(s) contain another loop inside them, creating O(n×m) complexity. With large collections this causes very long run times and heavy API throttling.`,
    recommendation: 'Refactor to avoid nested loops where possible. Use filter/select expressions, join operations, or break the inner work into a child flow called per iteration.',
    affectedItems: nestedLoops,
  };
}

const MESSAGING_OP_PATTERNS = [
  /send.*email/i, /post.*message/i, /post.*card/i, /send.*notification/i,
  /send.*sms/i, /send.*push/i, /send.*teams/i, /create.*item/i, /send.*adaptive/i,
];

function checkMessagingInLoop(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const loopsWithMessages: string[] = [];

  for (const [loopName, loopAction] of allActions) {
    if (loopAction.type?.toLowerCase() !== 'foreach') continue;
    const children = collectAllActions(loopAction.actions);
    const hasMessaging = children.some(([, a]) => {
      const op = (
        (a.inputs?.['operationId'] as string | undefined) ??
        (a.inputs?.['host'] as Record<string, unknown> | undefined)?.['operationId'] as string ?? ''
      ).toLowerCase();
      return MESSAGING_OP_PATTERNS.some((p) => p.test(op));
    });
    if (hasMessaging) loopsWithMessages.push(loopName);
  }

  if (loopsWithMessages.length === 0) return null;

  return {
    id: 'messaging-in-loop',
    severity: 'warning',
    title: 'Notification or messaging actions inside a loop',
    description: `${loopsWithMessages.length} loop(s) send emails, Teams messages, or notifications on every iteration. This can trigger throttling limits and flood recipients.`,
    recommendation: 'Accumulate results in an array or HTML table, then send a single summary message after the loop completes.',
    affectedItems: loopsWithMessages,
  };
}

function checkHttpRetryPolicy(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const httpNoRetry = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'http') return false;
    const retryPolicy = a.inputs?.['retryPolicy'] as Record<string, unknown> | undefined;
    // Default retry is 4 retries; warn only if explicitly set to none/zero
    return retryPolicy?.type === 'none' || retryPolicy?.count === 0;
  }).map(([name]) => name);

  if (httpNoRetry.length === 0) return null;

  return {
    id: 'http-retry-disabled',
    severity: 'warning',
    title: 'HTTP actions have retries disabled',
    description: `${httpNoRetry.length} HTTP action(s) have retry policy explicitly set to "None". Transient errors (429, 500, 502) will cause the flow to fail immediately without retry.`,
    recommendation: 'Use an exponential retry policy with at least 3 attempts for HTTP calls to external services to handle transient failures gracefully.',
    affectedItems: httpNoRetry,
  };
}

function checkFlowSize(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const count = allActions.length;
  if (count <= 50) return null;

  return {
    id: 'flow-too-large',
    severity: 'info',
    title: `Flow is very large (${count} actions)`,
    description: `This flow contains ${count} actions, making it difficult to read, test, and debug. Large monolithic flows are also harder to reuse.`,
    recommendation: 'Break the flow into smaller, focused child flows called via "Run a Child Flow". Aim for flows that do one thing well and stay under ~30 actions.',
  };
}

function checkSwitchWithoutDefault(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const switchesWithoutDefault = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'switch') return false;
    return !a.default || Object.keys(a.default.actions ?? {}).length === 0;
  }).map(([name]) => name);

  if (switchesWithoutDefault.length === 0) return null;

  return {
    id: 'switch-no-default',
    severity: 'warning',
    title: 'Switch actions missing a default case',
    description: `${switchesWithoutDefault.length} Switch action(s) have no default branch. Unexpected values will silently pass through without any handling.`,
    recommendation: 'Add a default branch to every Switch action to handle unexpected values — at minimum, log or terminate with a meaningful error message.',
    affectedItems: switchesWithoutDefault,
  };
}

function checkRecurrenceTimezone(triggers: Record<string, FlowTrigger>): AnalysisResult | null {
  const recurrenceWithoutTz = Object.entries(triggers).filter(([, t]) => {
    if (t.type?.toLowerCase() !== 'recurrence') return false;
    // startTime carries timezone info; if absent the trigger uses UTC implicitly
    return !(t as Record<string, unknown>)['startTime'];
  }).map(([name]) => name);

  if (recurrenceWithoutTz.length === 0) return null;

  return {
    id: 'recurrence-no-timezone',
    severity: 'info',
    title: 'Recurrence trigger has no explicit start time / time zone',
    description: 'Without a start time, the recurrence runs in UTC. This can cause flows to run at unexpected local times, especially across Daylight Saving Time changes.',
    recommendation: 'Set a start time on the recurrence trigger to anchor the schedule to a specific time zone. This prevents drift during DST transitions.',
    affectedItems: recurrenceWithoutTz,
  };
}

function checkDoUntilWithoutLimit(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  const unbounded = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'until') return false;
    const limit = a.limit as Record<string, unknown> | undefined;
    const count = limit?.count;
    // No count, or count is extremely high (>1000), is a risk
    return count === undefined || count === null || (typeof count === 'number' && count > 1000);
  }).map(([name]) => name);

  if (unbounded.length === 0) return null;

  return {
    id: 'do-until-no-limit',
    severity: 'warning',
    title: '"Do Until" loops without a meaningful iteration limit',
    description: `${unbounded.length} "Do Until" loop(s) have no iteration count limit (or an extremely high one). If the exit condition is never met, these loops will run until the flow times out after 30 days.`,
    recommendation: 'Always set a reasonable iteration limit (e.g., 60 for a polling loop with 1-minute delays). Combine with a Terminate action if the limit is hit to signal failure clearly.',
    affectedItems: unbounded,
  };
}

function checkHardcodedUrls(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  // HTTP actions where the URI is a static string with no expressions
  const hardcoded = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'http') return false;
    const uri = a.inputs?.['uri'] as string | undefined ?? a.inputs?.['url'] as string | undefined ?? '';
    // If the URI starts with http(s) and contains no expression syntax (@) it's fully hardcoded
    return /^https?:\/\//i.test(uri) && !uri.includes('@{') && !uri.includes('@variables') && !uri.includes('@parameters');
  }).map(([name]) => name);

  if (hardcoded.length === 0) return null;

  return {
    id: 'hardcoded-urls',
    severity: 'info',
    title: 'HTTP actions use hardcoded URLs',
    description: `${hardcoded.length} HTTP action(s) have fully static URLs with no expressions or parameters. These break when deploying to different environments (dev → test → prod).`,
    recommendation: 'Store environment-specific URLs in flow parameters or environment variables and reference them via expressions — making the flow portable across environments.',
    affectedItems: hardcoded.slice(0, 10),
  };
}

function checkDeepNesting(definition: FlowDefinition): AnalysisResult | null {
  const DEPTH_THRESHOLD = 4;
  const deepActions = collectActionsWithDepth(definition.actions)
    .filter(([, , depth]) => depth >= DEPTH_THRESHOLD)
    .map(([name]) => name);

  if (deepActions.length === 0) return null;

  return {
    id: 'deep-nesting',
    severity: 'info',
    title: 'Deeply nested actions detected',
    description: `${deepActions.length} action(s) are nested ${DEPTH_THRESHOLD}+ levels deep. Deep nesting makes flows very hard to read, test, and modify without introducing bugs.`,
    recommendation: 'Flatten logic by extracting deeply nested branches into child flows, or simplify conditions using filter/select expressions before the branch point.',
    affectedItems: deepActions.slice(0, 10),
  };
}

function checkTerminateWithoutFailedStatus(allActions: Array<[string, FlowAction]>): AnalysisResult | null {
  // Terminate actions that run after a failure but don't set runStatus to 'Failed'
  const badTerminations = allActions.filter(([, a]) => {
    if (a.type?.toLowerCase() !== 'terminate') return false;
    const ra = a.runAfter ?? {};
    const isInErrorPath = Object.values(ra).some((statuses) =>
      statuses.some((s) => ['Failed', 'TimedOut'].includes(s))
    );
    if (!isInErrorPath) return false;
    const runStatus = (a.inputs?.['runStatus'] as string | undefined ?? '').toLowerCase();
    return runStatus !== 'failed';
  }).map(([name]) => name);

  if (badTerminations.length === 0) return null;

  return {
    id: 'terminate-without-failed-status',
    severity: 'info',
    title: 'Error-path Terminate actions do not use "Failed" status',
    description: `${badTerminations.length} Terminate action(s) are in error paths but do not set run status to "Failed". This causes the flow run to appear as "Succeeded" even when an error occurred.`,
    recommendation: 'Set the Terminate action\'s run status to "Failed" with a descriptive error message when used in error-handling branches so failures are visible in run history and monitoring.',
    affectedItems: badTerminations,
  };
}

const DEFAULT_TRIGGER_NAME_PATTERNS = [
  /^manual(_\d+)?$/i,
  /^recurrence(_\d+)?$/i,
  /^Trigger(_\d+)?$/i,
  /^When_a_HTTP_request_is_received(_\d+)?$/i,
];

function checkTriggerDefaultName(triggers: Record<string, FlowTrigger>): AnalysisResult | null {
  const defaultNamed = Object.keys(triggers).filter((name) =>
    DEFAULT_TRIGGER_NAME_PATTERNS.some((p) => p.test(name))
  );
  if (defaultNamed.length === 0) return null;
  return {
    id: 'trigger-default-name',
    severity: 'info',
    title: 'Trigger uses a default name',
    description: `${defaultNamed.length} trigger(s) use auto-generated names, making the flow harder to understand at a glance.`,
    recommendation: 'Rename triggers with a descriptive name that explains what initiates the flow (e.g. "On new SharePoint item" instead of "manual").',
    affectedItems: defaultNamed,
  };
}

function checkUndocumentedActions(
  allActions: Array<[string, FlowAction]>,
  triggers: Record<string, FlowTrigger>,
): AnalysisResult | null {
  // Need at least 5 actions+triggers to make this meaningful
  const total = allActions.length + Object.keys(triggers).length;
  if (total < 5) return null;

  const undocumentedActions = allActions
    .filter(([, a]) => !a.description?.trim())
    .map(([name]) => name);

  const undocumentedTriggers = Object.entries(triggers)
    .filter(([, t]) => !t.description?.trim())
    .map(([name]) => name);

  const undocumented = [...undocumentedTriggers, ...undocumentedActions];
  const ratio = undocumented.length / total;

  // Only flag when 80%+ are undocumented
  if (ratio < 0.8) return null;

  return {
    id: 'no-action-descriptions',
    severity: 'info',
    title: 'Most steps have no description (comment)',
    description: `${undocumented.length} of ${total} actions and triggers have no description. Comments help future makers understand the purpose of each step without opening it.`,
    recommendation: 'Add a description to key actions — especially complex expressions, HTTP calls, variables, and any step whose purpose isn\'t obvious from its name.',
    affectedItems: undocumented.slice(0, 10),
  };
}



export function analyzeFlowDefinition(definition: FlowDefinition): AnalysisResult[] {
  const results: AnalysisResult[] = [];
  const triggers = definition.triggers ?? {};
  const allActions = collectAllActions(definition.actions);

  const checks = [
    // Original 10
    checkFlowDescription(definition),
    checkErrorHandling(allActions),
    checkTriggerConditions(triggers),
    checkHighFrequencyRecurrence(triggers),
    checkConcurrencyControl(triggers),
    checkDefaultActionNames(allActions),
    checkHttpWithoutTimeout(allActions),
    checkSensitiveDataExposure(allActions),
    checkForEachConcurrency(allActions),
    checkPaginationOnListActions(allActions),
    // New 10
    checkNestedForEach(allActions),
    checkMessagingInLoop(allActions),
    checkHttpRetryPolicy(allActions),
    checkFlowSize(allActions),
    checkSwitchWithoutDefault(allActions),
    checkRecurrenceTimezone(triggers),
    checkDoUntilWithoutLimit(allActions),
    checkHardcodedUrls(allActions),
    checkDeepNesting(definition),
    checkTerminateWithoutFailedStatus(allActions),
    // Naming & documentation
    checkTriggerDefaultName(triggers),
    checkUndocumentedActions(allActions, triggers),
  ];

  for (const r of checks) {
    if (r) results.push(r);
  }

  // Sort: critical → warning → info
  const order: Record<AnalysisSeverity, number> = { critical: 0, warning: 1, info: 2 };
  results.sort((a, b) => order[a.severity] - order[b.severity]);

  return results;
}
