"""
BPMN to SOP Parser - Full Implementation Following Guideline V2
Generates complete multi-paragraph structure with proper formatting
"""

import re
from lxml import etree
from typing import List, Dict, Tuple, Optional

class BPMNParser:
    """
    Complete BPMN parser implementing ALL Guideline V2 rules:
    - Multi-paragraph task format (title, desc)
    - Multi-paragraph gateway format (case title, explanation, routing)
    - Subprocess handling (no step number, special format)
    - Parallel gateways (AND) - "Proceed to Step X and Step Y"
    - Multi-inputs (parallel joins) - "Step Input: Step X and Step Y"
    - Intermediate events - "wait until [event] Then [action]"
    - Process ends - "Process Ends ([name])"
    """

    BPMN_NS = {'bpmn': 'http://www.omg.org/spec/BPMN/20100524/MODEL'}
    COMBINED_NS = {
        'bpmn': 'http://www.omg.org/spec/BPMN/20100524/MODEL',
        'zeebe': 'http://camunda.org/schema/zeebe/1.0'
    }
    BPMNDI_NS = {
        'bpmndi': 'http://www.omg.org/spec/BPMN/20100524/DI',
        'dc': 'http://www.omg.org/spec/DD/20100524/DC'
    }

    def __init__(self, xml_content: bytes):
        self.xml_content = xml_content
        self.root = etree.fromstring(xml_content)

        # Data structures
        self.tasks = {}
        self.gateways = {}
        self.flows = {}
        self.lanes = {}
        self.subprocesses = {}
        self.events = {}
        self.boundary_events = {}  # Track boundary events attached to tasks
        self.task_incoming = {}  # Track incoming flows for multi-input detection
        self.groups = {}        # Group elements with SLA
        self.shape_bounds = {}  # BPMNShape bounds for spatial queries
        self.lane_raci = {}     # Lane RACI data: lane_id -> {responsible, accountable, consulted, informed}
        self.element_to_lane_id = {}  # Element ID -> lane_id mapping

        self._parse_structure()

    def _parse_structure(self):
        """Parse complete BPMN structure"""

        # Parse lanes (including RACI documentation and element mapping)
        for lane in self.root.xpath('.//bpmn:lane', namespaces=self.BPMN_NS):
            lane_id = lane.get('id')
            lane_name = lane.get('name', '[LANE UNREADABLE]')
            self.lanes[lane_id] = lane_name

            # Parse RACI documentation from lane
            raci = {}
            for raci_type in ['responsible', 'accountable', 'consulted', 'informed']:
                raci_elements = lane.xpath(f'./bpmn:documentation[@textFormat="application/x-{raci_type}"]', namespaces=self.BPMN_NS)
                if raci_elements and raci_elements[0].text:
                    raci[raci_type] = raci_elements[0].text.strip()
                else:
                    raci[raci_type] = 'N/A'
            self.lane_raci[lane_id] = raci

            # Build element-to-lane mapping
            for ref in lane.xpath('./bpmn:flowNodeRef', namespaces=self.BPMN_NS):
                if ref.text:
                    self.element_to_lane_id[ref.text] = lane_id

        # Parse tasks (including all task types: standard, user, service, manual, script, call, send, receive, businessRule)
        for task in self.root.xpath('.//bpmn:task | .//bpmn:userTask | .//bpmn:serviceTask | .//bpmn:manualTask | .//bpmn:scriptTask | .//bpmn:callActivity | .//bpmn:sendTask | .//bpmn:receiveTask | .//bpmn:businessRuleTask', namespaces=self.BPMN_NS):
            task_id = task.get('id')
            task_name = task.get('name', '')
            lane_name = self._get_lane_for_element(task_id)
            step_number = self._extract_step_number(task_name)
            task_label = self._remove_step_number(task_name)

            incoming = [flow.text for flow in task.xpath('./bpmn:incoming', namespaces=self.BPMN_NS) if flow.text]
            outgoing = [flow.text for flow in task.xpath('./bpmn:outgoing', namespaces=self.BPMN_NS) if flow.text]

            # Parse documentation element for custom description (skip SLA docs)
            documentation = None
            doc_elements = task.xpath('./bpmn:documentation[not(@textFormat)]', namespaces=self.BPMN_NS)
            if doc_elements and doc_elements[0].text:
                documentation = doc_elements[0].text.strip()

            # Parse task-level SLA
            sla_elements = task.xpath('./bpmn:documentation[@textFormat="application/x-sla"]', namespaces=self.BPMN_NS)
            task_sla = sla_elements[0].text.strip() if sla_elements and sla_elements[0].text else None

            self.tasks[task_id] = {
                'name': task_name,
                'label': task_label,
                'lane': lane_name,
                'lane_id': self.element_to_lane_id.get(task_id),
                'number': step_number,
                'incoming': incoming,
                'outgoing': outgoing,
                'documentation': documentation,
                'sla': task_sla
            }

        # Parse gateways
        for gateway in self.root.xpath('.//bpmn:exclusiveGateway | .//bpmn:parallelGateway | .//bpmn:inclusiveGateway', namespaces=self.BPMN_NS):
            gateway_id = gateway.get('id')
            if 'parallel' in gateway.tag.lower():
                gateway_type = 'AND'
            elif 'inclusive' in gateway.tag.lower():
                gateway_type = 'OR'
            else:
                gateway_type = 'XOR'

            incoming = [flow.text for flow in gateway.xpath('./bpmn:incoming', namespaces=self.BPMN_NS) if flow.text]
            outgoing = [flow.text for flow in gateway.xpath('./bpmn:outgoing', namespaces=self.BPMN_NS) if flow.text]

            self.gateways[gateway_id] = {
                'type': gateway_type,
                'incoming': incoming,
                'outgoing': outgoing
            }

        # Parse sequence flows
        for flow in self.root.xpath('.//bpmn:sequenceFlow', namespaces=self.BPMN_NS):
            flow_id = flow.get('id')
            # Parse documentation element for sequence flow
            flow_documentation = None
            flow_doc_elements = flow.xpath('./bpmn:documentation', namespaces=self.BPMN_NS)
            if flow_doc_elements and flow_doc_elements[0].text:
                flow_documentation = flow_doc_elements[0].text.strip()
            self.flows[flow_id] = {
                'source': flow.get('sourceRef'),
                'target': flow.get('targetRef'),
                'name': flow.get('name', ''),
                'documentation': flow_documentation
            }

        # Parse subprocesses
        for subprocess in self.root.xpath('.//bpmn:subProcess', namespaces=self.BPMN_NS):
            subprocess_id = subprocess.get('id')
            subprocess_name = subprocess.get('name', 'Unnamed Process')
            # Clean up newlines in subprocess name
            subprocess_name = re.sub(r'[\n\r\t]+', ' ', subprocess_name)
            subprocess_name = re.sub(r'\s+', ' ', subprocess_name).strip()

            self.subprocesses[subprocess_id] = {
                'name': subprocess_name,
                'incoming': [flow.text for flow in subprocess.xpath('./bpmn:incoming', namespaces=self.BPMN_NS) if flow.text],
                'outgoing': [flow.text for flow in subprocess.xpath('./bpmn:outgoing', namespaces=self.BPMN_NS) if flow.text]
            }

        # Parse events (start, end, intermediate)
        for event in self.root.xpath('.//bpmn:startEvent | .//bpmn:endEvent | .//bpmn:intermediateCatchEvent | .//bpmn:intermediateThrowEvent', namespaces=self.BPMN_NS):
            event_id = event.get('id')
            event_name = event.get('name', '')
            event_type = 'start' if 'start' in event.tag.lower() else ('end' if 'end' in event.tag.lower() else 'intermediate')

            self.events[event_id] = {
                'name': event_name,
                'type': event_type,
                'incoming': [flow.text for flow in event.xpath('./bpmn:incoming', namespaces=self.BPMN_NS) if flow.text],
                'outgoing': [flow.text for flow in event.xpath('./bpmn:outgoing', namespaces=self.BPMN_NS) if flow.text]
            }

        # Parse boundary events (attached to tasks/subprocesses)
        for boundary in self.root.xpath('.//bpmn:boundaryEvent', namespaces=self.BPMN_NS):
            boundary_id = boundary.get('id')
            boundary_name = boundary.get('name', '')
            attached_to = boundary.get('attachedToRef')  # The task/subprocess this is attached to
            cancel_activity = boundary.get('cancelActivity', 'true').lower() == 'true'  # interrupting or not

            # Detect event type (timer, message, signal, error, etc.)
            event_type = None
            if boundary.xpath('./bpmn:timerEventDefinition', namespaces=self.BPMN_NS):
                event_type = 'timer'
            elif boundary.xpath('./bpmn:messageEventDefinition', namespaces=self.BPMN_NS):
                event_type = 'message'
            elif boundary.xpath('./bpmn:signalEventDefinition', namespaces=self.BPMN_NS):
                event_type = 'signal'
            elif boundary.xpath('./bpmn:errorEventDefinition', namespaces=self.BPMN_NS):
                event_type = 'error'

            if attached_to:
                if attached_to not in self.boundary_events:
                    self.boundary_events[attached_to] = []

                self.boundary_events[attached_to].append({
                    'id': boundary_id,
                    'name': boundary_name,
                    'interrupting': cancel_activity,
                    'event_type': event_type,
                    'outgoing': [flow.text for flow in boundary.xpath('./bpmn:outgoing', namespaces=self.BPMN_NS) if flow.text]
                })

        # Parse groups with SLA documentation
        for group in self.root.xpath('.//bpmn:group', namespaces=self.BPMN_NS):
            group_id = group.get('id')
            sla_doc = group.xpath('./bpmn:documentation[@textFormat="application/x-sla"]', namespaces=self.BPMN_NS)
            if sla_doc and sla_doc[0].text:
                self.groups[group_id] = {
                    'sla': sla_doc[0].text.strip()
                }

        # Parse BPMNShape bounds for spatial queries (group membership)
        for shape in self.root.xpath('.//bpmndi:BPMNShape', namespaces=self.BPMNDI_NS):
            element_ref = shape.get('bpmnElement')
            bounds = shape.find('{http://www.omg.org/spec/DD/20100524/DC}Bounds')
            if bounds is not None and element_ref:
                self.shape_bounds[element_ref] = {
                    'x': float(bounds.get('x', 0)),
                    'y': float(bounds.get('y', 0)),
                    'width': float(bounds.get('width', 0)),
                    'height': float(bounds.get('height', 0))
                }

    def _get_lane_for_element(self, element_id: str) -> str:
        """Get lane name for element and clean special Unicode characters"""
        for lane in self.root.xpath('.//bpmn:lane', namespaces=self.BPMN_NS):
            flow_node_refs = lane.xpath('./bpmn:flowNodeRef', namespaces=self.BPMN_NS)
            for ref in flow_node_refs:
                if ref.text == element_id:
                    lane_name = lane.get('name', '[LANE UNREADABLE]')
                    # Remove Arabic diacritics and other special Unicode marks (combining characters)
                    import unicodedata
                    lane_name = ''.join(c for c in lane_name if not unicodedata.combining(c))
                    return lane_name
        return '[LANE UNREADABLE]'

    def _get_task_sla(self, task_id: str) -> Tuple[Optional[str], Optional[str]]:
        """Get SLA for a task (from task-level or group-level)
        Returns: (sla_value, sla_group_id) or (None, None)
        """
        task = self.tasks.get(task_id)
        if not task:
            return None, None

        # Check task-level SLA first (takes priority)
        if task.get('sla'):
            return task['sla'], None

        # Check if task is in any SLA group (spatial containment)
        task_bounds = self.shape_bounds.get(task_id)
        if not task_bounds:
            return None, None

        task_cx = task_bounds['x'] + task_bounds['width'] / 2
        task_cy = task_bounds['y'] + task_bounds['height'] / 2

        for group_id, group_data in self.groups.items():
            group_bounds = self.shape_bounds.get(group_id)
            if not group_bounds:
                continue

            gx = group_bounds['x']
            gy = group_bounds['y']
            gw = group_bounds['width']
            gh = group_bounds['height']

            if gx <= task_cx <= gx + gw and gy <= task_cy <= gy + gh:
                return group_data['sla'], group_id

        return None, None

    def _extract_step_number(self, task_name: str) -> Optional[str]:
        """Extract step number from task name"""
        match = re.match(r'^\s*(\d+)\s*[.:\-]?\s*', task_name)
        return match.group(1) if match else None

    def _remove_step_number(self, task_name: str) -> str:
        """Remove step number from task name and clean up newlines"""
        # Replace newline characters (including &#10; decoded as \n) with spaces
        cleaned_name = re.sub(r'[\n\r\t]+', ' ', task_name)
        # Remove step number prefix
        cleaned_name = re.sub(r'^\s*\d+\s*[.:\-]?\s*', '', cleaned_name)
        # Clean up multiple spaces and strip
        cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
        return cleaned_name

    def _get_start_event_numbers(self) -> Dict[str, int]:
        """
        Map start event IDs to their input numbers (Input 1, Input 2, etc.)

        Ordered by which task/step they first connect to:
        - Triggers connecting to Step 1 are numbered first (Input 1, Input 2, etc.)
        - Triggers connecting to Step 2 are numbered next, and so on
        - Within same step, ordered by XML parse order

        This ensures consistent numbering like:
        - Input 1 = B2B Request (connects to Task 1)
        - Input 2 = Customer Call (connects to Task 2)
        - Input 3 = Zabbix Error (connects to Task 2)
        """
        # Collect start events with the step number of the first task they connect to
        start_events_with_step = []

        for event_id, event_data in self.events.items():
            if event_data['type'] == 'start':
                # Find the first task this trigger connects to (trace through gateways if needed)
                first_step = self._find_first_task_from_element(event_id)
                # Use a high number if no task found (puts them at end)
                step_num = int(first_step) if first_step else 9999
                start_events_with_step.append((event_id, step_num))

        # Sort by step number (primary), then by original order (secondary - for stability)
        start_events_with_step.sort(key=lambda x: x[1])

        # Assign numbers
        numbers = {}
        for n, (event_id, _) in enumerate(start_events_with_step, start=1):
            numbers[event_id] = n

        return numbers

    def _find_first_task_from_element(self, element_id: str, visited: set = None) -> Optional[str]:
        """
        Find the first numbered task that an element connects to.
        Traces forward through gateways and intermediate events.
        Returns the step number or None.
        """
        if visited is None:
            visited = set()

        if element_id in visited:
            return None
        visited.add(element_id)

        # Get outgoing flows from this element
        outgoing_flows = []
        if element_id in self.events:
            outgoing_flows = self.events[element_id].get('outgoing', [])
        elif element_id in self.gateways:
            outgoing_flows = self.gateways[element_id].get('outgoing', [])

        # Check each outgoing flow
        for flow_id in outgoing_flows:
            if flow_id not in self.flows:
                continue

            target_id = self.flows[flow_id]['target']

            # If target is a task, return its step number
            if target_id in self.tasks:
                step_num = self.tasks[target_id].get('number')
                if step_num:
                    return step_num

            # If target is a gateway or intermediate event, continue tracing
            if target_id in self.gateways or target_id in self.events:
                result = self._find_first_task_from_element(target_id, visited)
                if result:
                    return result

        return None

    def _get_target_step_numbers(self, element_id: str) -> List[Tuple[str, str, str]]:
        """
        Get list of (target_type, target_step_or_name, flow_name) tuples
        target_type: 'task', 'gateway', 'subprocess', 'end', 'intermediate'
        """
        targets = []

        # Get outgoing flows
        if element_id in self.tasks:
            flow_ids = self.tasks[element_id]['outgoing']
        elif element_id in self.gateways:
            flow_ids = self.gateways[element_id]['outgoing']
        elif element_id in self.subprocesses:
            flow_ids = self.subprocesses[element_id]['outgoing']
        elif element_id in self.events:
            flow_ids = self.events[element_id]['outgoing']
        else:
            return targets

        for flow_id in flow_ids:
            if flow_id not in self.flows:
                continue

            flow = self.flows[flow_id]
            target_id = flow['target']
            flow_name = flow['name']

            if target_id in self.tasks:
                target_step = self.tasks[target_id]['number']
                if target_step:
                    targets.append(('task', target_step, flow_name))
            elif target_id in self.gateways:
                targets.append(('gateway', target_id, flow_name))
            elif target_id in self.subprocesses:
                targets.append(('subprocess', target_id, flow_name))
            elif target_id in self.events:
                if self.events[target_id]['type'] == 'end':
                    end_name = self.events[target_id]['name'] or 'Process Complete'
                    targets.append(('end', end_name, flow_name))
                elif self.events[target_id]['type'] == 'intermediate':
                    targets.append(('intermediate', target_id, flow_name))

        return targets

    def _trace_gateway_to_task(self, gateway_id: str, visited: set = None) -> Optional[str]:
        """
        Trace back through a gateway to find the source task's step number.
        Used for parallel join detection when one path goes through intermediate gateways.
        Returns the step number of the first task found, or None.
        """
        if visited is None:
            visited = set()

        if gateway_id in visited:
            return None
        visited.add(gateway_id)

        if gateway_id not in self.gateways:
            return None

        gateway = self.gateways[gateway_id]
        incoming_flows = gateway.get('incoming', [])

        for flow_id in incoming_flows:
            if flow_id not in self.flows:
                continue

            source_id = self.flows[flow_id]['source']

            # If source is a task, return its step number
            if source_id in self.tasks and self.tasks[source_id]['number']:
                return self.tasks[source_id]['number']

            # If source is another gateway, trace through it
            if source_id in self.gateways:
                result = self._trace_gateway_to_task(source_id, visited)
                if result:
                    return result

        return None

    def _trace_back_to_split_gateway(self, element_id: str, visited: set = None) -> Optional[tuple]:
        """
        Trace back through the flow graph to find the originating split gateway.
        Returns (gateway_id, gateway_type) or None.
        A split gateway has 1 incoming flow and multiple outgoing flows.
        """
        if visited is None:
            visited = set()

        # Avoid infinite loops
        if element_id in visited:
            return None
        visited.add(element_id)

        # Check if this element is a gateway
        if element_id in self.gateways:
            gw = self.gateways[element_id]
            # Check if it's a split (1 input, multiple outputs)
            if len(gw.get('incoming', [])) == 1 and len(gw.get('outgoing', [])) > 1:
                return (element_id, gw['type'])
            # If it's a join or merge, continue tracing back
            incoming_flows = gw.get('incoming', [])
        elif element_id in self.tasks:
            incoming_flows = self.tasks[element_id].get('incoming', [])
        else:
            return None

        # If no incoming flows, we've reached the start
        if not incoming_flows:
            return None

        # If multiple incoming flows, this is a join point - stop here
        # (We don't want to trace back through joins)
        if len(incoming_flows) > 1:
            return None

        # Trace back through the single incoming flow
        flow_id = incoming_flows[0]
        if flow_id not in self.flows:
            return None

        source_id = self.flows[flow_id]['source']
        return self._trace_back_to_split_gateway(source_id, visited)

    def _detect_multi_input(self, task_id: str) -> Optional[str]:
        """
        Detect if task is a multi-input (parallel join or inclusive join)
        Returns formatted string like:
        - "Step Input: Step 23 and Step 26" (parallel/AND join)
        - "Step Input: Step 40 and/or Step 41 and/or Step 42" (inclusive/OR join)
        - "Step Input: Step 1 or Input 2" (step + trigger combination)

        NOTE: If all incoming flows trace back to the same XOR gateway split,
        this is NOT a multi-input - only one path executes.

        IMPORTANT: If there's a mix of parallel/inclusive joins AND XOR branches,
        only show the Step Input from the parallel/inclusive joins (ignore XOR branches).
        """
        if task_id not in self.tasks:
            return None

        incoming_flows = self.tasks[task_id]['incoming']
        # Note: Don't return None for len=1, could be from a join gateway!
        if len(incoming_flows) == 0:
            return None

        # Check if any incoming flow comes directly from a parallel/inclusive join gateway
        has_parallel_or_inclusive_join = False
        parallel_or_inclusive_sources = []
        xor_sources = []

        for flow_id in incoming_flows:
            if flow_id not in self.flows:
                continue

            source_id = self.flows[flow_id]['source']

            # Check if the direct source is a gateway
            if source_id in self.gateways:
                source_gw = self.gateways[source_id]
                if source_gw['type'] in ['AND', 'OR']:
                    # This is a parallel or inclusive join
                    has_parallel_or_inclusive_join = True
                    parallel_or_inclusive_sources.append((flow_id, source_id, source_gw['type']))
                elif source_gw['type'] == 'XOR':
                    xor_sources.append((flow_id, source_id))
            else:
                # Source is a task or other element - trace back
                if source_id in self.tasks:
                    split_info = self._trace_back_to_split_gateway(source_id)
                    if split_info and split_info[1] == 'XOR':
                        xor_sources.append((flow_id, source_id))

        # If there are parallel/inclusive joins, use ONLY those for Step Input (ignore XOR branches)
        if has_parallel_or_inclusive_join:
            source_steps = []
            join_type = 'and'  # Default to AND

            for flow_id, gateway_id, gateway_type in parallel_or_inclusive_sources:
                # Determine join type based on gateway
                if gateway_type == 'OR':
                    join_type = 'or'
                elif gateway_type == 'AND':
                    join_type = 'and'

                # Get the tasks before the gateway (trace through intermediate gateways if needed)
                gateway_incoming = self.gateways[gateway_id]['incoming']
                for gw_flow_id in gateway_incoming:
                    if gw_flow_id in self.flows:
                        gw_source_id = self.flows[gw_flow_id]['source']
                        if gw_source_id in self.tasks and self.tasks[gw_source_id]['number']:
                            source_steps.append(self.tasks[gw_source_id]['number'])
                        elif gw_source_id in self.gateways:
                            # Source is another gateway - trace back to find the task
                            task_step = self._trace_gateway_to_task(gw_source_id)
                            if task_step:
                                source_steps.append(task_step)

            if len(source_steps) > 1:
                source_steps = sorted(set(source_steps), key=int)  # Remove duplicates and sort
                connector = ' and/or Step ' if join_type == 'or' else ' and Step '
                return f"Step Input: Step {connector.join(source_steps)}"

            # If parallel/inclusive join doesn't have multiple task sources,
            # check for step + trigger combination (task + start event)
            return self._detect_step_trigger_input(task_id)

        # No parallel/inclusive joins found - proceed with old logic for other cases
        # Trace back each incoming flow to find its originating split gateway
        split_gateways = []
        has_xor_source = False

        for flow_id in incoming_flows:
            if flow_id not in self.flows:
                continue

            source_id = self.flows[flow_id]['source']

            # Check if the direct source is an XOR gateway
            if source_id in self.gateways:
                source_gw = self.gateways[source_id]
                if source_gw['type'] == 'XOR':
                    has_xor_source = True

            # Trace back whether source is a task or gateway
            if source_id in self.tasks or source_id in self.gateways:
                split_info = self._trace_back_to_split_gateway(source_id)
                if split_info:
                    split_gateways.append(split_info)
                    if split_info[1] == 'XOR':
                        has_xor_source = True
                else:
                    # Couldn't trace back to a split - treat as normal multi-input
                    split_gateways.append(None)
            else:
                # Source is neither task nor gateway
                split_gateways.append(None)

        # If ANY of the incoming flows involves an XOR gateway, this is NOT a true multi-input
        # XOR means only one path executes at a time (could be loops or conditional flows)
        if has_xor_source:
            # But first check for Step + Trigger combination before returning None
            return self._detect_step_trigger_input(task_id)

        # Check if all paths trace back to XOR splits
        if len(split_gateways) == len(incoming_flows) and len(split_gateways) > 0:
            # Filter out None values
            non_none_splits = [s for s in split_gateways if s is not None]

            if len(non_none_splits) == len(split_gateways):
                # All flows traced back to a split
                # Check if they're ALL XOR splits (same or different)
                all_xor = all(gw_type == 'XOR' for gw_id, gw_type in non_none_splits)

                if all_xor:
                    # All flows come from XOR splits - NOT a multi-input
                    # Only one path executes in XOR, so these can't all arrive simultaneously
                    # But first check for Step + Trigger combination
                    return self._detect_step_trigger_input(task_id)

        # Otherwise, proceed with normal multi-input detection
        source_steps = []
        join_type = 'and'  # Default to AND

        for flow_id in incoming_flows:
            if flow_id in self.flows:
                source_id = self.flows[flow_id]['source']

                # Check if source is a task
                if source_id in self.tasks and self.tasks[source_id]['number']:
                    source_steps.append(self.tasks[source_id]['number'])
                # Check if source is a gateway (joining after gateway split)
                elif source_id in self.gateways:
                    gateway_type = self.gateways[source_id]['type']

                    # Determine join type based on gateway
                    if gateway_type == 'OR':
                        join_type = 'or'
                    elif gateway_type == 'AND':
                        join_type = 'and'

                    # Get the tasks before the gateway
                    gateway_incoming = self.gateways[source_id]['incoming']
                    for gw_flow_id in gateway_incoming:
                        if gw_flow_id in self.flows:
                            gw_source_id = self.flows[gw_flow_id]['source']
                            if gw_source_id in self.tasks and self.tasks[gw_source_id]['number']:
                                source_steps.append(self.tasks[gw_source_id]['number'])

        if len(source_steps) > 1:
            source_steps = sorted(set(source_steps), key=int)  # Remove duplicates and sort
            connector = ' and/or Step ' if join_type == 'or' else ' and Step '
            steps_str = connector.join([f"{s}" for s in source_steps])
            return f"Step Input: Step {steps_str}"

        # Finally, check for Step + Trigger combination
        return self._detect_step_trigger_input(task_id)

    def _detect_step_trigger_input(self, task_id: str) -> Optional[str]:
        """
        Detect if task has inputs from BOTH previous steps AND triggers (start events).

        According to Guideline V2 Multi Inputs section:
        - If a task has only ONE type of input (just step OR just trigger) -> NO Step Input text
        - If a task has inputs from BOTH a step AND a trigger -> "Step Input: Step X or Input Y"

        Example: If task 2 receives input from Step 1 AND from trigger "BB" (Input 2):
        -> "Y Step Input: Step 1 or Input 2"

        Triggers are start events, numbered as Input 1, Input 2, etc. based on their order.
        A trigger reaches a task "directly" if it connects without passing through another numbered task.

        IMPORTANT: Reverts (flows from higher-numbered steps back to lower-numbered steps)
        are NOT counted as step sources. Only FORWARD flows (from earlier steps) count.
        This is because reverts are re-executions, not initial entry points.
        """
        if task_id not in self.tasks:
            return None

        incoming_flows = self.tasks[task_id]['incoming']
        if len(incoming_flows) == 0:
            return None

        # Get current task's step number to filter out reverts
        current_task_num_str = self.tasks[task_id].get('number')
        if not current_task_num_str:
            return None
        current_task_num = int(current_task_num_str)

        start_event_numbers = self._get_start_event_numbers()

        step_sources = set()     # Step numbers that feed this task (forward flows only)
        trigger_sources = set()  # Input numbers (from start events) that feed this task

        def trace_source(element_id: str, visited: set):
            """
            Recursively trace backward to find task or start event sources.
            Stops at tasks (step sources) and start events (trigger sources).
            Continues through gateways and intermediate events.

            For step sources, only counts FORWARD flows (step_num < current_task_num).
            Reverts (step_num >= current_task_num) are excluded.
            """
            if element_id in visited:
                return
            visited.add(element_id)

            # If it's a task, record as step source (if forward flow) and stop
            if element_id in self.tasks:
                num = self.tasks[element_id].get('number')
                if num:
                    source_step_num = int(num)
                    # Only count as step source if it's a FORWARD flow (earlier step)
                    # Reverts (from later steps) are NOT initial entry points
                    if source_step_num < current_task_num:
                        step_sources.add(num)
                return

            # If it's a start event, record as trigger source and stop
            if element_id in self.events:
                if self.events[element_id]['type'] == 'start':
                    if element_id in start_event_numbers:
                        trigger_sources.add(start_event_numbers[element_id])
                    return
                # For intermediate events, continue tracing backward
                for flow_id in self.events[element_id].get('incoming', []):
                    if flow_id in self.flows:
                        trace_source(self.flows[flow_id]['source'], visited)
                return

            # If it's a gateway, continue tracing backward through all incoming flows
            if element_id in self.gateways:
                for flow_id in self.gateways[element_id].get('incoming', []):
                    if flow_id in self.flows:
                        trace_source(self.flows[flow_id]['source'], visited)
                return

            # If it's a subprocess, continue tracing backward
            if element_id in self.subprocesses:
                for flow_id in self.subprocesses[element_id].get('incoming', []):
                    if flow_id in self.flows:
                        trace_source(self.flows[flow_id]['source'], visited)
                return

        # Trace from each incoming flow of the task
        for flow_id in incoming_flows:
            if flow_id in self.flows:
                source_id = self.flows[flow_id]['source']
                trace_source(source_id, set())

        # Only create Step Input text if we have BOTH steps AND triggers
        if step_sources and trigger_sources:
            parts = []
            # Add step sources (sorted numerically)
            for step_num in sorted(step_sources, key=int):
                parts.append(f"Step {step_num}")
            # Add trigger sources (sorted numerically)
            for input_num in sorted(trigger_sources):
                parts.append(f"Input {input_num}")
            return f"Step Input: {' or '.join(parts)}"

        return None

    def _check_intermediate_event(self, task_id: str) -> Optional[str]:
        """
        Check if there's an intermediate event before this task
        Returns event description or None
        """
        if task_id not in self.tasks:
            return None

        incoming_flows = self.tasks[task_id]['incoming']

        for flow_id in incoming_flows:
            if flow_id in self.flows:
                source_id = self.flows[flow_id]['source']
                if source_id in self.events and self.events[source_id]['type'] == 'intermediate':
                    event_name = self.events[source_id]['name']
                    if event_name:
                        return event_name

        return None

    def _check_boundary_events(self, task_id: str) -> List[str]:
        """
        Check for boundary events attached to this task
        Returns list of formatted boundary event texts

        According to Guideline V2:
        - Interrupting: "If [condition], stop the activity and proceed to step X"
        - Non-Interrupting: "If [condition], proceed to step X and complete the activity, then proceed to step Y"

        For condition formatting based on event type:
        - Timer events (timerEventDefinition): "If performing the activity took more than [time], ..."
        - Other events (message, signal, error): "If [event name] during performing the activity, ..."
        """
        if task_id not in self.boundary_events:
            return []

        boundary_texts = []

        for boundary in self.boundary_events[task_id]:
            event_name = boundary['name']
            event_type = boundary.get('event_type')
            if not event_name:
                continue

            # Find where the boundary event leads
            target_step = None
            if boundary['outgoing']:
                flow_id = boundary['outgoing'][0]
                if flow_id in self.flows:
                    target_id = self.flows[flow_id]['target']
                    if target_id in self.tasks:
                        target_step = self.tasks[target_id]['number']

            if not target_step:
                continue

            # Format the condition part based on event type
            if event_type == 'timer':
                # Timer event: "If performing the activity took more than 30 min"
                condition = f"performing the activity took more than {event_name}"
            else:
                # Message, signal, error, or other events: "If [event] during performing the activity"
                condition = f"{event_name} during performing the activity"

            # Format according to interrupting or non-interrupting
            if boundary['interrupting']:
                # Interrupting: "If [condition], stop the activity and proceed to step X"
                boundary_text = f"If {condition}, stop the activity and proceed to step {target_step}"
            else:
                # Non-Interrupting: "If [condition], proceed to step X and complete the activity, then proceed to step Y"
                # For non-interrupting, we need to find the normal next step (Y)
                # This requires checking the task's normal outgoing flow
                normal_next_step = None
                normal_next_is_end = False
                end_event_name = None

                if task_id in self.tasks:
                    task_targets = self._get_target_step_numbers(task_id)
                    for target_type, target_value, flow_name in task_targets:
                        if target_type == 'task':
                            normal_next_step = target_value
                            break
                        elif target_type == 'gateway':
                            # Follow through gateway to find next task or end
                            gateway_id = target_value
                            gateway_targets = self._get_target_step_numbers(gateway_id)
                            for gw_target_type, gw_target_value, gw_flow_name in gateway_targets:
                                if gw_target_type == 'task':
                                    normal_next_step = gw_target_value
                                    break
                                elif gw_target_type == 'end':
                                    normal_next_is_end = True
                                    end_event_name = gw_target_value
                                    break
                            break
                        elif target_type == 'end':
                            normal_next_is_end = True
                            end_event_name = target_value
                            break

                if normal_next_step:
                    boundary_text = f"If {condition}, proceed to step {target_step} and complete the activity, then proceed to step {normal_next_step}"
                elif normal_next_is_end:
                    boundary_text = f"If {condition}, proceed to step {target_step} and complete the activity, then Process Ends ({end_event_name})"
                else:
                    # Fallback if normal next step not found
                    boundary_text = f"If {condition}, proceed to step {target_step} and complete the activity"

            boundary_texts.append(boundary_text)

        return boundary_texts

    def _check_intermediate_before_subprocess(self, subprocess_id: str) -> Optional[str]:
        """
        Check if there's an intermediate event before this subprocess
        Returns event name or None
        """
        if subprocess_id not in self.subprocesses:
            return None

        incoming_flows = self.subprocesses[subprocess_id]['incoming']

        for flow_id in incoming_flows:
            if flow_id in self.flows:
                source_id = self.flows[flow_id]['source']
                if source_id in self.events and self.events[source_id]['type'] == 'intermediate':
                    event_name = self.events[source_id]['name']
                    if event_name:
                        return event_name

        return None

    def _check_task_intermediate_chain(self, task_id: str, current_step_num: str = None) -> Optional[str]:
        """
        Check if this task leads to intermediate event → (subprocess OR gateway OR task)

        Patterns handled:
        1. Task → Intermediate Event → Subprocess → Task
           Returns: "Wait for [event] and Then Start [subprocess] Process Then Proceed/Revert to Step X"

        2. Task → Intermediate Event → Gateway → Tasks
           Returns: "Wait for [event] and Then Proceed to Step X and Step Y"

        3. Task → Intermediate Event → Task
           Returns: "Wait for [event] and Then Proceed/Revert to Step X"

        This implements the guideline rule (page 16-17): If task leads to intermediate event,
        write the wait text IN the task
        """
        # Get what follows this task
        targets = self._get_target_step_numbers(task_id)
        if not targets or targets[0][0] != 'intermediate':
            return None

        # Task leads to an intermediate event
        event_id = targets[0][1]
        event = self.events.get(event_id)
        if not event:
            return None

        event_name = event['name']

        # Check if event_name already starts with "Wait for" or "wait for"
        # to avoid duplication like "Wait for Wait for time"
        if event_name.lower().startswith('wait for '):
            wait_prefix = ""
            event_text = event_name
        elif event_name.lower().startswith('wait until '):
            # Also check for "Wait until" to avoid "Wait for Wait until..."
            wait_prefix = ""
            event_text = event_name
        else:
            wait_prefix = "Wait for "
            event_text = event_name

        # Check what follows the intermediate event
        event_targets = self._get_target_step_numbers(event_id)
        if not event_targets:
            return f"{wait_prefix}{event_text}"

        first_target_type = event_targets[0][0]
        first_target_value = event_targets[0][1]

        # Case 1: Intermediate → Subprocess
        if first_target_type == 'subprocess':
            subprocess_id = first_target_value
            subprocess = self.subprocesses.get(subprocess_id)
            if not subprocess:
                return f"{wait_prefix}{event_text}"

            subprocess_name = subprocess['name']
            process_suffix = "" if subprocess_name.lower().endswith("process") else " Process"

            # Check what follows the subprocess
            subprocess_targets = self._get_target_step_numbers(subprocess_id)
            if subprocess_targets and subprocess_targets[0][0] == 'task':
                next_step = subprocess_targets[0][1]
                # Check if it's a revert (going back to earlier step) or proceed
                routing_verb = "Proceed"
                if current_step_num and int(next_step) <= int(current_step_num):
                    routing_verb = "Revert"
                return f"{wait_prefix}{event_text} and Then Start {subprocess_name}{process_suffix} Then {routing_verb} to Step {next_step}"
            elif subprocess_targets and subprocess_targets[0][0] == 'end':
                end_name = subprocess_targets[0][1]
                return f"{wait_prefix}{event_text} and Then Start {subprocess_name}{process_suffix}, then Process Ends ({end_name})"
            else:
                return f"{wait_prefix}{event_text} and Then Start {subprocess_name}{process_suffix}"

        # Case 2: Intermediate → Gateway
        elif first_target_type == 'gateway':
            gateway_id = first_target_value
            gateway = self.gateways.get(gateway_id)
            if not gateway:
                return f"{wait_prefix}{event_text}"

            # Get targets from the gateway
            gateway_targets = self._get_target_step_numbers(gateway_id)

            # Check if it's a parallel (AND) or inclusive (OR) gateway split
            if gateway['type'] in ['AND', 'OR'] and len(gateway.get('outgoing', [])) > 1:
                # Collect task targets
                task_targets = sorted([t[1] for t in gateway_targets if t[0] == 'task'], key=int)

                if task_targets:
                    connector = ' and/or Step ' if gateway['type'] == 'OR' else ' and Step '
                    steps_str = connector.join(task_targets)
                    return f"{wait_prefix}{event_text} and Then Proceed to Step {steps_str}"

            # XOR gateway or other - just pick first task target
            for gw_target_type, gw_target_value, _ in gateway_targets:
                if gw_target_type == 'task':
                    routing_verb = "Proceed"
                    if current_step_num and int(gw_target_value) <= int(current_step_num):
                        routing_verb = "Revert"
                    return f"{wait_prefix}{event_text} and Then {routing_verb} to Step {gw_target_value}"
                elif gw_target_type == 'end':
                    return f"{wait_prefix}{event_text}, then Process Ends ({gw_target_value})"

            return f"{wait_prefix}{event_text}"

        # Case 3: Intermediate → Task (direct)
        elif first_target_type == 'task':
            next_step = first_target_value
            routing_verb = "Proceed"
            if current_step_num and int(next_step) <= int(current_step_num):
                routing_verb = "Revert"
            return f"{wait_prefix}{event_text} and Then {routing_verb} to Step {next_step}"

        # Case 4: Intermediate → End
        elif first_target_type == 'end':
            end_name = first_target_value
            return f"{wait_prefix}{event_text}, then Process Ends ({end_name})"

        return f"{wait_prefix}{event_text}"

    def _check_task_intermediate_subprocess_chain(self, task_id: str, current_step_num: str = None) -> Optional[str]:
        """
        DEPRECATED: Use _check_task_intermediate_chain instead.
        Keeping for backwards compatibility, now just calls the new method.
        """
        return self._check_task_intermediate_chain(task_id, current_step_num)

    def generate_sop_rows(self) -> List[Dict]:
        """
        Generate complete SOP rows with multi-paragraph structure

        Returns list of dicts with structure:
        {
            'ref': '1' or '2A',
            'is_gateway': False/True,
            'paragraphs': [
                {'text': 'Submit Request', 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'},
                {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},  # empty
                {'text': 'The Department shall...', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'}
            ]
        }
        """
        sop_rows = []

        # Sort tasks by step number
        sorted_tasks = sorted(
            [(tid, t) for tid, t in self.tasks.items() if t['number'] is not None],
            key=lambda x: int(x[1]['number'])
        )

        for task_id, task in sorted_tasks:
            step_num = task['number']
            lane_name = task['lane']
            task_label = task['label']

            # Check if this task leads to intermediate event → subprocess chain
            task_subprocess_chain = self._check_task_intermediate_subprocess_chain(task_id, step_num)

            # Check for intermediate event (only if not part of subprocess chain)
            intermediate_event = None if task_subprocess_chain else self._check_intermediate_event(task_id)

            # Check for multiple inputs (parallel join or inclusive join)
            multi_input = self._detect_multi_input(task_id)

            # Build title - append "Step Input:" if multiple incoming flows
            title_text = task_label
            if multi_input:
                title_text = f"{title_text} {multi_input}"

            # Build description
            lane_suffix = ""

            # Get documentation from task (custom description from BPMN)
            documentation = task.get('documentation')

            # Split documentation into lines (preserve line breaks)
            doc_lines = []
            extra_doc_lines = []
            if documentation:
                # Split on newlines, strip each line but keep non-empty ones
                raw_lines = documentation.split('\n')
                doc_lines = [line.rstrip() for line in raw_lines]
                # First line is used for desc_text, rest are extra paragraphs
                if doc_lines:
                    first_line = doc_lines[0].strip()
                    extra_doc_lines = doc_lines[1:]  # Keep indentation for extra lines

            # Helper to ensure text ends with a period
            def ensure_period(text):
                text = text.rstrip()
                if text and not text.endswith(('.', '!', '?')):
                    text += '.'
                return text

            # Build task description (always include this)
            if intermediate_event:
                # Intermediate event format: "shall wait until [event] Then [action]"
                if doc_lines and doc_lines[0].strip():
                    first_line = doc_lines[0].strip()
                    # Use documentation, check if it starts with "shall"
                    if first_line.lower().startswith('shall '):
                        desc_text = f"The {lane_name}{lane_suffix} shall wait until {intermediate_event} Then {first_line[6:].strip()}"
                    else:
                        desc_text = f"The {lane_name}{lane_suffix} shall wait until {intermediate_event} Then {first_line}"
                else:
                    desc_text = f"The {lane_name}{lane_suffix} shall wait until {intermediate_event} Then {task_label.lower()}"
                desc_text = ensure_period(desc_text)
            else:
                if doc_lines and doc_lines[0].strip():
                    first_line = doc_lines[0].strip()
                    # Use documentation, check if it starts with "shall"
                    if first_line.lower().startswith('shall '):
                        # Already has "shall", use as-is
                        desc_text = f"The {lane_name}{lane_suffix} {first_line}"
                    else:
                        # Add "shall" before the documentation
                        desc_text = f"The {lane_name}{lane_suffix} shall {first_line}"
                else:
                    # No documentation, use task label as before
                    desc_text = f"The {lane_name}{lane_suffix} shall {task_label.lower()}"
                desc_text = ensure_period(desc_text)

            # Check for boundary events attached to this task
            boundary_events = self._check_boundary_events(task_id)

            # Check what follows this task (need to do this early to detect backward flows)
            targets = self._get_target_step_numbers(task_id)

            # Check if there's a backward flow (revert to earlier step)
            revert_step = None
            for target_type, target_value, flow_name in targets:
                if target_type == 'task':
                    target_num = int(target_value)
                    current_num = int(step_num)
                    if target_num < current_num:
                        revert_step = target_value
                        break

            # Create task row with proper paragraph structure
            # Structure: title (with step input if applicable), empty, description, optional extra doc lines, optional subprocess chain, optional boundary events, optional revert
            paragraphs = [
                {'text': title_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'},
                {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},  # empty
                {'text': desc_text, 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'}
            ]

            # Add extra documentation lines if present (multi-line documentation from BPMN)
            if extra_doc_lines:
                # Add empty line before extra doc lines
                paragraphs.append({'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'})
                for extra_line in extra_doc_lines:
                    # Keep each line (preserve indentation), ensure last line has period
                    line_text = extra_line.rstrip()
                    if line_text:  # Only add non-empty lines
                        paragraphs.append({'text': line_text, 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'})
                # Ensure the last paragraph ends with a period
                if paragraphs and paragraphs[-1]['text']:
                    paragraphs[-1]['text'] = ensure_period(paragraphs[-1]['text'])

            # Add subprocess chain if present (comes after description, before boundary events)
            if task_subprocess_chain:
                paragraphs.extend([
                    {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},  # empty
                    {'text': task_subprocess_chain, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                ])

            # Add boundary events if present
            if boundary_events:
                for boundary_text in boundary_events:
                    paragraphs.extend([
                        {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},  # empty
                        {'text': boundary_text, 'font_size': 11, 'bold': True, 'alignment': 'JUSTIFY'}
                    ])

            # Add revert flow if present
            if revert_step:
                paragraphs.extend([
                    {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},  # empty
                    {'text': f'Revert to Step {revert_step}', 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                ])

            sla_value, sla_group = self._get_task_sla(task_id)

            # Get RACI from lane
            lane_id = task.get('lane_id')
            task_raci = self.lane_raci.get(lane_id, {'responsible': 'N/A', 'accountable': 'N/A', 'consulted': 'N/A', 'informed': 'N/A'})

            sop_rows.append({
                'ref': step_num,
                'is_gateway': False,
                'sla': sla_value,
                'sla_group': sla_group,
                'raci': task_raci,
                'paragraphs': paragraphs
            })

            if len(targets) == 0:
                continue

            # Check if this task feeds into a parallel/inclusive JOIN gateway going to end event
            # When parallel branches converge and go directly to end (no task after join),
            # each branch task says "Wait until step [other] completed then Process Ends ([name])"
            parallel_join_end_handled = False
            seen_join_gateways = set()
            for target_type_j, target_value_j, flow_name_j in targets:
                if target_type_j == 'gateway' and target_value_j not in seen_join_gateways:
                    seen_join_gateways.add(target_value_j)
                    gw_j = self.gateways.get(target_value_j)
                    if gw_j and gw_j['type'] in ['AND', 'OR'] and len(gw_j.get('incoming', [])) > 1:
                        # This is a parallel/inclusive JOIN
                        join_targets = self._get_target_step_numbers(target_value_j)
                        end_targets = [jt for jt in join_targets if jt[0] == 'end']
                        task_targets_after = [jt for jt in join_targets if jt[0] == 'task']

                        # Only handle if join goes to END event (not to a task)
                        if end_targets and not task_targets_after:
                            end_name = end_targets[0][1]
                            # Find other tasks feeding into this join
                            other_steps = []
                            for gw_flow_id in gw_j['incoming']:
                                if gw_flow_id in self.flows:
                                    source_id = self.flows[gw_flow_id]['source']
                                    if source_id == task_id:
                                        continue
                                    if source_id in self.tasks:
                                        other_step = self.tasks[source_id].get('number')
                                        if other_step:
                                            other_steps.append(other_step)
                                    elif source_id in self.gateways:
                                        traced = self._trace_gateway_to_task(source_id)
                                        if traced:
                                            other_steps.append(traced)

                            if other_steps:
                                other_steps = sorted(set(other_steps), key=int)
                                steps_str = ' and step '.join(other_steps)
                                wait_text = f"Wait until step {steps_str} completed then Process Ends ({end_name})"
                                sop_rows[-1]['paragraphs'].extend([
                                    {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                                    {'text': wait_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                                ])
                                parallel_join_end_handled = True
                            break

            if parallel_join_end_handled:
                continue

            # Check if next element is a gateway
            if len(targets) == 1 and targets[0][0] == 'gateway':
                gateway_id = targets[0][1]
                gateway = self.gateways[gateway_id]

                # Check if it's a parallel (AND) or inclusive (OR) gateway SPLIT
                if (gateway['type'] in ['AND', 'OR']) and len(gateway['incoming']) == 1 and len(gateway['outgoing']) > 1:
                    # Get all targets from the gateway
                    gateway_targets = self._get_target_step_numbers(gateway_id)

                    # Filter task targets and sort them
                    task_targets = sorted([t[1] for t in gateway_targets if t[0] == 'task'], key=int)

                    if len(task_targets) > 1:
                        # Parallel/Inclusive split to multiple steps
                        # AND gateway: "Proceed to Step X and Step Y"
                        # OR gateway: "Proceed to Step X and/or Step Y and/or Step Z"
                        connector = ' and/or Step ' if gateway['type'] == 'OR' else ' and Step '
                        steps_str = connector.join(task_targets)
                        proceed_text = f"Proceed to Step {steps_str}"

                        # Add extra paragraphs to the last task row
                        sop_rows[-1]['paragraphs'].extend([
                            {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                            {'text': proceed_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                        ])
                    # If not multiple task targets, it's a parallel/inclusive JOIN - just continue
                elif gateway['type'] == 'XOR':
                    # ONLY XOR (exclusive) gateways show as cases
                    gateway_rows = self._generate_gateway_rows(step_num, gateway_id, task_raci)
                    sop_rows.extend(gateway_rows)
                # For any other gateway type not handled above, just continue

            # Check if next element is a subprocess
            elif len(targets) == 1 and targets[0][0] == 'subprocess':
                subprocess_id = targets[0][1]

                # Check if there's an intermediate event before the subprocess
                intermediate_event = self._check_intermediate_before_subprocess(subprocess_id)

                subprocess = self.subprocesses[subprocess_id]
                process_name = subprocess['name']

                # Find what comes after subprocess
                subprocess_targets = self._get_target_step_numbers(subprocess_id)

                # Check if process name already ends with "Process" to avoid duplication
                process_suffix = "" if process_name.lower().endswith("process") else " Process"

                # Check what comes after the subprocess
                if intermediate_event:
                    # With intermediate event before subprocess
                    # Format: "Wait for [event] and Then Start [subprocess] Then Proceed/Revert to Step X"
                    if subprocess_targets and subprocess_targets[0][0] == 'task':
                        next_step = subprocess_targets[0][1]
                        # Check if it's a revert (going back to earlier step) or proceed
                        routing_verb = "Revert" if int(next_step) <= int(step_num) else "Proceed"
                        subprocess_text = f"Wait for {intermediate_event} and Then Start {process_name}{process_suffix} Then {routing_verb} to Step {next_step}"
                    elif subprocess_targets and subprocess_targets[0][0] == 'end':
                        # Subprocess goes to end event
                        end_event_name = subprocess_targets[0][1]
                        subprocess_text = f"Wait for {intermediate_event} and Then Start {process_name}{process_suffix}, then Process Ends ({end_event_name})"
                    else:
                        subprocess_text = f"Wait for {intermediate_event} and Then Start {process_name}{process_suffix}"
                else:
                    # No intermediate event
                    # Check if subprocess leads to end event or next task
                    if subprocess_targets and subprocess_targets[0][0] == 'end':
                        # Subprocess ends the process
                        # Format: "Start [subprocess], then Process Ends ([end event name])" per Guideline page 12 & Point Dismantling Step 13
                        end_event_name = subprocess_targets[0][1]
                        subprocess_text = f"Start {process_name}{process_suffix}, then Process Ends ({end_event_name})"
                    elif subprocess_targets and subprocess_targets[0][0] == 'task':
                        # Subprocess goes to next task
                        # Format: "Start [subprocess] Then Proceed/Revert to Step X" per Guideline V2 page 12
                        next_step = subprocess_targets[0][1]
                        # Check if it's a revert (going back to earlier step) or proceed
                        routing_verb = "Revert" if int(next_step) <= int(step_num) else "Proceed"
                        subprocess_text = f"Start {process_name}{process_suffix} Then {routing_verb} to Step {next_step}"
                    else:
                        # Fallback
                        subprocess_text = f"Start {process_name}{process_suffix}"

                # Always add subprocess to current task paragraphs (not as separate row)
                sop_rows[-1]['paragraphs'].extend([
                    {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                    {'text': subprocess_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                ])

            # Check if next element is an end event
            elif len(targets) == 1 and targets[0][0] == 'end':
                end_event_name = targets[0][1]
                # Add "Process Ends (end event name)" to the last task row
                sop_rows[-1]['paragraphs'].extend([
                    {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                    {'text': f'Process Ends ({end_event_name})', 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                ])

            # Check for parallel gateway (AND) - multiple target tasks
            elif len(targets) > 1 and all(t[0] == 'task' for t in targets):
                # Parallel split - "Proceed to Step X and Step Y"
                target_steps = sorted([t[1] for t in targets], key=int)
                steps_str = ' and Step '.join(target_steps)
                proceed_text = f"Proceed to Step {steps_str}"

                paragraphs = [
                    {'text': proceed_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                ]

                # Add as additional paragraph to last row (no new ref)
                sop_rows[-1]['paragraphs'].extend([
                    {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                    {'text': proceed_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
                ])

        return sop_rows

    def _generate_subprocess_row(self, subprocess_id: str) -> Optional[Dict]:
        """
        Generate subprocess row (NO step number!)
        Format: "Start [Process Name] Process Then Proceed to Step X"
        """
        if subprocess_id not in self.subprocesses:
            return None

        subprocess = self.subprocesses[subprocess_id]
        process_name = subprocess['name']

        # Find what comes after subprocess
        targets = self._get_target_step_numbers(subprocess_id)

        if targets and targets[0][0] == 'task':
            next_step = targets[0][1]
            text = f"Start: {process_name} Process Then Proceed to Step {next_step}"
        else:
            text = f"Start: {process_name} Process"

        paragraphs = [
            {'text': text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
        ]

        return {
            'ref': '',  # NO step number for subprocess!
            'is_gateway': False,
            'paragraphs': paragraphs
        }

    def _generate_gateway_rows(self, parent_step: str, gateway_id: str, parent_raci: Dict = None) -> List[Dict]:
        """
        Generate gateway case rows with proper lettering and formatting
        5 paragraphs per case: Title, empty, explanation, empty, routing
        """
        if gateway_id not in self.gateways:
            return []

        gateway_cases = []
        gateway = self.gateways[gateway_id]
        targets = self._get_target_step_numbers(gateway_id)

        if not targets:
            return []

        # Build flow documentation lookup: map flow_name to flow documentation
        flow_doc_by_name = {}
        for flow_id in gateway['outgoing']:
            if flow_id in self.flows:
                f = self.flows[flow_id]
                if f.get('documentation') and f.get('name'):
                    flow_doc_by_name[f['name']] = f['documentation']

        # Build cases with sorting info
        cases_with_dest = []
        for target_type, target_value, flow_name in targets:
            if target_type == 'task':
                target_num = int(target_value)
                parent_num = int(parent_step)

                # Self-loops (Step 6 → Step 6) are reverts (redo the step)
                # Reverts are: target_num <= parent_num
                is_revert = target_num <= parent_num

                # Sorting key:
                # - Ends: (-1, 0) - comes first (Case A)
                # - Reverts: (0, target_num) - next, farthest back first
                # - Proceeds: (1, -target_num) - last, farthest forward first
                if is_revert:
                    sort_key = (0, target_num)  # Reverts come after ends, farthest back first
                else:
                    sort_key = (1, -target_num)  # Proceeds come last, farthest forward first

                cases_with_dest.append({
                    'flow_name': flow_name if flow_name else '[CONDITION UNLABELED]',
                    'flow_documentation': flow_doc_by_name.get(flow_name),
                    'target_step': target_value,
                    'target_num': target_num,
                    'is_revert': is_revert,
                    'sort_key': sort_key,
                    'target_type': 'task'
                })
            elif target_type == 'end':
                cases_with_dest.append({
                    'flow_name': flow_name if flow_name else 'Complete',
                    'flow_documentation': flow_doc_by_name.get(flow_name),
                    'target_step': None,
                    'end_name': target_value,
                    'is_revert': False,
                    'sort_key': (-1, 0),  # Ends come first (Case A)
                    'target_type': 'end'
                })
            elif target_type == 'subprocess':
                # Subprocess after gateway
                subprocess = self.subprocesses.get(target_value)
                if subprocess:
                    cases_with_dest.append({
                        'flow_name': flow_name if flow_name else 'Proceed',
                        'flow_documentation': flow_doc_by_name.get(flow_name),
                        'subprocess_name': subprocess['name'],
                        'subprocess_id': target_value,  # Store the actual subprocess ID
                        'is_revert': False,
                        'sort_key': (1, 5000),  # Treat as proceed, but after normal proceeds
                        'target_type': 'subprocess'
                    })
            elif target_type == 'gateway':
                # Gateway leads to another gateway - trace through it
                intermediate_gateway_id = target_value
                intermediate_gateway = self.gateways.get(intermediate_gateway_id)
                intermediate_targets = self._get_target_step_numbers(intermediate_gateway_id)

                # Check if intermediate gateway is an AND/OR split OR join-then-split
                # Split: 1 input, multiple outputs
                # Join-then-split: multiple inputs, multiple outputs (treat as split for routing purposes)
                is_parallel_split = (intermediate_gateway and
                                   intermediate_gateway['type'] in ['AND', 'OR'] and
                                   len(intermediate_gateway.get('outgoing', [])) > 1)

                if is_parallel_split:
                    # Collect all task targets
                    task_targets = []
                    for inter_type, inter_value, inter_flow_name in intermediate_targets:
                        if inter_type == 'task':
                            task_targets.append(inter_value)

                    if task_targets:
                        # Create a single case with parallel routing
                        task_targets_sorted = sorted(task_targets, key=int)
                        parent_num = int(parent_step)
                        first_target_num = int(task_targets_sorted[0])
                        is_revert = first_target_num < parent_num

                        if is_revert:
                            sort_key = (0, first_target_num)  # Reverts after ends
                        else:
                            sort_key = (1, -first_target_num)  # Proceeds last

                        # Create parallel/inclusive routing text
                        connector = ' and/or Step ' if intermediate_gateway['type'] == 'OR' else ' and Step '
                        steps_str = connector.join(task_targets_sorted)

                        cases_with_dest.append({
                            'flow_name': flow_name if flow_name else '[CONDITION UNLABELED]',
                            'flow_documentation': flow_doc_by_name.get(flow_name),
                            'parallel_targets': task_targets_sorted,
                            'gateway_type': intermediate_gateway['type'],
                            'is_revert': is_revert,
                            'sort_key': sort_key,
                            'target_type': 'parallel'
                        })
                else:
                    # Not a parallel/inclusive split - process targets individually
                    for inter_type, inter_value, inter_flow_name in intermediate_targets:
                        if inter_type == 'task':
                            target_num = int(inter_value)
                            parent_num = int(parent_step)
                            is_revert = target_num < parent_num

                            if is_revert:
                                sort_key = (0, target_num)
                            else:
                                sort_key = (1, -target_num)

                            cases_with_dest.append({
                                'flow_name': flow_name if flow_name else '[CONDITION UNLABELED]',
                                'flow_documentation': flow_doc_by_name.get(flow_name),
                                'target_step': inter_value,
                                'target_num': target_num,
                                'is_revert': is_revert,
                                'sort_key': sort_key,
                                'target_type': 'task'
                            })
                        elif inter_type == 'end':
                            cases_with_dest.append({
                                'flow_name': flow_name if flow_name else 'Complete',
                                'flow_documentation': flow_doc_by_name.get(flow_name),
                                'target_step': None,
                                'end_name': inter_value,
                                'is_revert': False,
                                'sort_key': (-1, 0),  # Ends come first
                                'target_type': 'end'
                            })
            elif target_type == 'intermediate':
                # Gateway → Intermediate Event → [target]
                # Trace through the intermediate event to find the actual destination
                event = self.events.get(target_value)
                if not event:
                    continue

                event_name = event['name'] or 'Event'
                event_targets = self._get_target_step_numbers(target_value)

                if not event_targets:
                    continue

                first_et_type = event_targets[0][0]
                first_et_value = event_targets[0][1]

                if first_et_type == 'task':
                    target_num = int(first_et_value)
                    parent_num = int(parent_step)
                    is_revert = target_num <= parent_num

                    sort_key = (0, target_num) if is_revert else (1, -target_num)

                    cases_with_dest.append({
                        'flow_name': flow_name if flow_name else '[CONDITION UNLABELED]',
                        'flow_documentation': flow_doc_by_name.get(flow_name),
                        'target_step': first_et_value,
                        'target_num': target_num,
                        'is_revert': is_revert,
                        'sort_key': sort_key,
                        'event_name': event_name,
                        'target_type': 'intermediate_task'
                    })
                elif first_et_type == 'end':
                    cases_with_dest.append({
                        'flow_name': flow_name if flow_name else 'Complete',
                        'flow_documentation': flow_doc_by_name.get(flow_name),
                        'target_step': None,
                        'end_name': first_et_value,
                        'is_revert': False,
                        'sort_key': (-1, 0),  # Ends come first
                        'event_name': event_name,
                        'target_type': 'intermediate_end'
                    })
                elif first_et_type == 'subprocess':
                    subprocess = self.subprocesses.get(first_et_value)
                    if subprocess:
                        cases_with_dest.append({
                            'flow_name': flow_name if flow_name else 'Proceed',
                            'flow_documentation': flow_doc_by_name.get(flow_name),
                            'subprocess_name': subprocess['name'],
                            'subprocess_id': first_et_value,
                            'is_revert': False,
                            'sort_key': (1, 5000),
                            'event_name': event_name,
                            'target_type': 'intermediate_subprocess'
                        })
                elif first_et_type == 'gateway':
                    # Intermediate → Gateway - trace further
                    inter_gw_id = first_et_value
                    inter_gw = self.gateways.get(inter_gw_id)
                    inter_gw_targets = self._get_target_step_numbers(inter_gw_id)

                    if inter_gw and inter_gw['type'] in ['AND', 'OR'] and len(inter_gw.get('outgoing', [])) > 1:
                        task_targets = sorted([t[1] for t in inter_gw_targets if t[0] == 'task'], key=int)
                        if task_targets:
                            parent_num = int(parent_step)
                            first_target_num = int(task_targets[0])
                            is_revert = first_target_num < parent_num
                            sort_key = (0, first_target_num) if is_revert else (1, -first_target_num)

                            cases_with_dest.append({
                                'flow_name': flow_name if flow_name else '[CONDITION UNLABELED]',
                                'flow_documentation': flow_doc_by_name.get(flow_name),
                                'parallel_targets': task_targets,
                                'gateway_type': inter_gw['type'],
                                'is_revert': is_revert,
                                'sort_key': sort_key,
                                'event_name': event_name,
                                'target_type': 'intermediate_parallel'
                            })
                    else:
                        for ig_type, ig_value, _ in (inter_gw_targets or []):
                            if ig_type == 'task':
                                target_num = int(ig_value)
                                parent_num = int(parent_step)
                                is_revert = target_num < parent_num
                                sort_key = (0, target_num) if is_revert else (1, -target_num)

                                cases_with_dest.append({
                                    'flow_name': flow_name if flow_name else '[CONDITION UNLABELED]',
                                    'flow_documentation': flow_doc_by_name.get(flow_name),
                                    'target_step': ig_value,
                                    'target_num': target_num,
                                    'is_revert': is_revert,
                                    'sort_key': sort_key,
                                    'event_name': event_name,
                                    'target_type': 'intermediate_task'
                                })
                                break
                            elif ig_type == 'end':
                                cases_with_dest.append({
                                    'flow_name': flow_name if flow_name else 'Complete',
                                    'flow_documentation': flow_doc_by_name.get(flow_name),
                                    'target_step': None,
                                    'end_name': ig_value,
                                    'is_revert': False,
                                    'sort_key': (-1, 0),  # Ends come first
                                    'event_name': event_name,
                                    'target_type': 'intermediate_end'
                                })
                                break

        # Sort by sort_key: ends first, then reverts (farthest back first), then proceeds (farthest forward first)
        cases_with_dest.sort(key=lambda x: x['sort_key'])

        # Assign letters
        for i, case in enumerate(cases_with_dest):
            letter = chr(ord('A') + i)
            case_title = f"Case {letter}: {case['flow_name']}"

            # Determine routing
            if case['target_type'] == 'task':
                routing_verb = "Revert to" if case['is_revert'] else "Proceed to"
                routing_text = f"{routing_verb} Step {case['target_step']}"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'end':
                routing_text = f"Process Ends ({case['end_name']})"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'subprocess':
                # Use the stored subprocess ID directly
                subprocess_id = case.get('subprocess_id')

                # Get next step after subprocess
                subprocess_targets = self._get_target_step_numbers(subprocess_id) if subprocess_id else []

                # Check if subprocess leads to a task, gateway, or end event
                if subprocess_targets and subprocess_targets[0][0] == 'task':
                    next_step = subprocess_targets[0][1]
                    # Check if it's a revert (going back to earlier step) or proceed
                    routing_verb = "Revert" if int(next_step) <= int(parent_step) else "Proceed"
                    routing_text = f"Start {case['subprocess_name']} Process, then {routing_verb} to Step {next_step}"
                elif subprocess_targets and subprocess_targets[0][0] == 'gateway':
                    # Subprocess leads to a gateway - check if it's an AND/OR split
                    sp_gw_id = subprocess_targets[0][1]
                    sp_gw = self.gateways.get(sp_gw_id)
                    if sp_gw and sp_gw['type'] in ['AND', 'OR'] and len(sp_gw.get('outgoing', [])) > 1:
                        # Get task targets from the gateway
                        sp_gw_targets = self._get_target_step_numbers(sp_gw_id)
                        task_targets = sorted([t[1] for t in sp_gw_targets if t[0] == 'task'], key=int)
                        if task_targets:
                            connector = ' and/or Step ' if sp_gw['type'] == 'OR' else ' and Step '
                            steps_str = connector.join(task_targets)
                            # Check if any target is a revert
                            first_target = int(task_targets[0])
                            routing_verb = "Revert" if first_target <= int(parent_step) else "Proceed"
                            routing_text = f"Start {case['subprocess_name']} Process, then {routing_verb} to Step {steps_str}"
                        else:
                            routing_text = f"Start {case['subprocess_name']} Process"
                    else:
                        routing_text = f"Start {case['subprocess_name']} Process"
                elif subprocess_targets and subprocess_targets[0][0] == 'end':
                    end_event_name = subprocess_targets[0][1]
                    routing_text = f"Start {case['subprocess_name']} Process, then Process Ends ({end_event_name})"
                else:
                    routing_text = f"Start {case['subprocess_name']} Process"

                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'parallel':
                # Parallel/inclusive routing (e.g., "Proceed to Step 27 and Step 28")
                parallel_targets = case['parallel_targets']
                gateway_type = case['gateway_type']
                routing_verb = "Revert to" if case['is_revert'] else "Proceed to"
                connector = ' and/or Step ' if gateway_type == 'OR' else ' and Step '
                steps_str = connector.join(parallel_targets)
                routing_text = f"{routing_verb} Step {steps_str}"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'intermediate_task':
                # Gateway → Intermediate Event → Task
                event_name = case['event_name']
                routing_verb = "Revert to" if case['is_revert'] else "Proceed to"
                routing_text = f"Wait until {event_name} Then {routing_verb} Step {case['target_step']}"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'intermediate_end':
                # Gateway → Intermediate Event → End
                event_name = case['event_name']
                routing_text = f"Wait until {event_name}, then Process Ends ({case['end_name']})"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'intermediate_subprocess':
                # Gateway → Intermediate Event → Subprocess
                event_name = case['event_name']
                subprocess_name = case['subprocess_name']
                subprocess_id = case.get('subprocess_id')
                process_suffix = "" if subprocess_name.lower().endswith("process") else " Process"
                subprocess_targets = self._get_target_step_numbers(subprocess_id) if subprocess_id else []

                if subprocess_targets and subprocess_targets[0][0] == 'task':
                    next_step = subprocess_targets[0][1]
                    rv = "Revert" if int(next_step) <= int(parent_step) else "Proceed"
                    routing_text = f"Wait until {event_name} Then Start {subprocess_name}{process_suffix}, then {rv} to Step {next_step}"
                elif subprocess_targets and subprocess_targets[0][0] == 'end':
                    end_name = subprocess_targets[0][1]
                    routing_text = f"Wait until {event_name} Then Start {subprocess_name}{process_suffix}, then Process Ends ({end_name})"
                else:
                    routing_text = f"Wait until {event_name} Then Start {subprocess_name}{process_suffix}"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"
            elif case['target_type'] == 'intermediate_parallel':
                # Gateway → Intermediate Event → Parallel/Inclusive Gateway → Tasks
                event_name = case['event_name']
                parallel_targets = case['parallel_targets']
                gateway_type = case['gateway_type']
                routing_verb = "Revert to" if case['is_revert'] else "Proceed to"
                connector = ' and/or Step ' if gateway_type == 'OR' else ' and Step '
                steps_str = connector.join(parallel_targets)
                routing_text = f"Wait until {event_name} Then {routing_verb} Step {steps_str}"
                explanation_text = case.get('flow_documentation') or f"[Condition explanation for {case['flow_name']}]"

            # Build 5 paragraphs
            paragraphs = [
                {'text': case_title, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'},
                {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                {'text': explanation_text, 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                {'text': '', 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'},
                {'text': routing_text, 'font_size': 12, 'bold': True, 'alignment': 'JUSTIFY'}
            ]

            gateway_cases.append({
                'ref': f"{parent_step}{letter}",
                'is_gateway': True,
                'sla': None,
                'sla_group': None,
                'raci': parent_raci or {'responsible': 'N/A', 'accountable': 'N/A', 'consulted': 'N/A', 'informed': 'N/A'},
                'paragraphs': paragraphs
            })

        return gateway_cases

    def get_process_inputs(self) -> str:
        """Extract start events as process inputs, ordered by which task they connect to"""
        # Use the same ordering as _get_start_event_numbers() for consistency
        start_event_numbers = self._get_start_event_numbers()

        # Build list of (input_number, name) tuples
        inputs_with_numbers = []
        for event_id, input_num in start_event_numbers.items():
            event_data = self.events[event_id]
            # First try to get the start event name
            if event_data['name']:
                inputs_with_numbers.append((input_num, event_data['name']))
            else:
                # If start event has no name, check the outgoing sequence flow name
                outgoing_flows = event_data.get('outgoing', [])
                for flow_id in outgoing_flows:
                    if flow_id in self.flows and self.flows[flow_id]['name']:
                        inputs_with_numbers.append((input_num, self.flows[flow_id]['name']))
                        break

        # Sort by input number and format
        inputs_with_numbers.sort(key=lambda x: x[0])

        if inputs_with_numbers:
            return '\n'.join([f"{num}. {name}" for num, name in inputs_with_numbers])
        return ''

    def get_process_outputs(self) -> str:
        """Extract end events as process outputs"""
        outputs = []
        for event_id, event_data in self.events.items():
            if event_data['type'] == 'end' and event_data['name']:
                outputs.append(event_data['name'])

        if outputs:
            return '\n'.join([f"{i+1}. {out}" for i, out in enumerate(outputs)])
        return ''

    def extract_bpmn_metadata(self) -> dict:
        """
        Extract metadata from BPMN elements for form auto-population.

        Extracts:
        - process_name: from <bpmn:participant name="...">
        - purpose: from <bpmn:participant>/<bpmn:documentation>
        - scope: from <bpmn:process>/<bpmn:documentation textFormat="application/x-scope">
        - process_code: from <zeebe:versionTag value="...">
        - abbreviations_list: from <zeebe:properties>/<zeebe:property name="..." value="...">

        Returns dict with keys matching the form field names.
        Only includes keys where data was found in the BPMN.
        """
        metadata = {}

        # 1. Participant name -> process_name
        participant = self.root.xpath('.//bpmn:participant[@name]', namespaces=self.BPMN_NS)
        if participant:
            name = participant[0].get('name', '').strip()
            if name:
                metadata['process_name'] = name

        # 2. Participant documentation -> purpose
        participant_doc = self.root.xpath(
            './/bpmn:participant/bpmn:documentation',
            namespaces=self.BPMN_NS
        )
        if participant_doc and participant_doc[0].text:
            metadata['purpose'] = participant_doc[0].text.strip()

        # 3. Process documentation with textFormat="application/x-scope" -> scope
        scope_doc = self.root.xpath(
            './/bpmn:process/bpmn:documentation[@textFormat="application/x-scope"]',
            namespaces=self.BPMN_NS
        )
        if scope_doc and scope_doc[0].text:
            metadata['scope'] = scope_doc[0].text.strip()

        # 4. zeebe:versionTag -> process_code
        version_tag = self.root.xpath(
            './/bpmn:process/bpmn:extensionElements/zeebe:versionTag',
            namespaces=self.COMBINED_NS
        )
        if version_tag:
            value = version_tag[0].get('value', '').strip()
            if value:
                metadata['process_code'] = value

        # 5. zeebe:properties -> abbreviations_list
        zeebe_properties = self.root.xpath(
            './/bpmn:process/bpmn:extensionElements/zeebe:properties/zeebe:property',
            namespaces=self.COMBINED_NS
        )
        if zeebe_properties:
            abbreviations = []
            for prop in zeebe_properties:
                term = prop.get('name', '').strip()
                definition = prop.get('value', '').strip()
                if term or definition:
                    abbreviations.append({'term': term, 'definition': definition})
            if abbreviations:
                metadata['abbreviations_list'] = abbreviations

        # 6. Lane names -> references_approvals (for auto-generating "N/A | Lane Approval" rows)
        lane_names = []
        for lane in self.root.xpath('.//bpmn:lane[@name]', namespaces=self.BPMN_NS):
            name = lane.get('name', '').strip()
            if name:
                lane_names.append(name)
        if lane_names:
            metadata['lane_names'] = lane_names

        # 7. Process documentation with textFormat="application/x-policy" -> general_policies_list
        policy_docs = self.root.xpath(
            './/bpmn:process/bpmn:documentation[@textFormat="application/x-policy"]',
            namespaces=self.BPMN_NS
        )
        if policy_docs:
            policies = []
            for idx, policy_doc in enumerate(policy_docs, start=1):
                if policy_doc.text:
                    policies.append({
                        'ref': str(idx),
                        'policy': policy_doc.text.strip()
                    })
            if policies:
                metadata['general_policies_list'] = policies

        return metadata


def extract_metadata_from_bpmn(xml_content: bytes) -> dict:
    """
    Extract metadata from BPMN file for form auto-population.
    Does NOT generate SOP rows - just extracts metadata fields.
    """
    try:
        parser = BPMNParser(xml_content)
        metadata = parser.extract_bpmn_metadata()

        # Also include auto-populated inputs/outputs
        metadata['inputs'] = parser.get_process_inputs()
        metadata['outputs'] = parser.get_process_outputs()

        return metadata
    except Exception as e:
        print(f"[ERROR] BPMN Metadata Extraction Failed: {e}")
        return {}


def parse_bpmn_to_sop(xml_content: bytes, metadata: dict) -> dict:
    """
    Main parsing function
    Returns context with structured SOP rows for template
    """
    try:
        parser = BPMNParser(xml_content)
        sop_rows = parser.generate_sop_rows()

        # Auto-populate inputs and outputs from BPMN if not provided
        auto_inputs = parser.get_process_inputs()
        auto_outputs = parser.get_process_outputs()

        context = {
            'process_name': metadata.get('process_name', ''),
            'process_code': metadata.get('process_code', ''),
            'issued_by': metadata.get('issued_by', 'Business Excellence'),  # Default value
            'release_date': metadata.get('release_date', 'TBD'),
            'process_owner': metadata.get('process_owner', ''),
            'purpose': metadata.get('purpose', ''),
            'scope': metadata.get('scope', ''),
            'abbreviations': metadata.get('abbreviations', ''),
            'references': metadata.get('references', ''),
            'inputs': metadata.get('inputs', auto_inputs),  # Use auto-populated if not provided
            'outputs': metadata.get('outputs', auto_outputs),  # Use auto-populated if not provided
            'abbreviations_list': metadata.get('abbreviations_list', []),
            'references_list': metadata.get('references_list', []),
            'general_policies_list': metadata.get('general_policies_list', []),
            'steps': sop_rows
        }

        return context

    except Exception as e:
        print(f"[ERROR] BPMN Parsing Failed: {e}")
        import traceback
        traceback.print_exc()

        return {
            'process_name': metadata.get('process_name', 'ERROR'),
            'process_code': '',
            'issued_by': '',
            'release_date': 'TBD',
            'purpose': '',
            'scope': '',
            'abbreviations': '',
            'references': '',
            'inputs': '',
            'outputs': '',
            'steps': [{
                'ref': 'ERROR',
                'is_gateway': False,
                'paragraphs': [{'text': f"Parse error: {str(e)}", 'font_size': 11, 'bold': False, 'alignment': 'JUSTIFY'}]
            }]
        }
