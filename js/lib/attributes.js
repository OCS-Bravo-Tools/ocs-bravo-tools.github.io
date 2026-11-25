/**
 * Questions, scale and the 30 attributes, in the order of the original sheet
 * (rows D23:D52). Pure data, no DOM.
 *
 * The one-line definitions are drafts written to fit one phone line. They must
 * be checked against ADP 6-22 before an official evaluation.
 */

/** The four Yes/No questions, in the sheet's own wording. */
export const QUESTIONS = [
  { id: 'lead', text: 'Would you trust this OC to lead soldiers?' },
  { id: 'need', text: 'Could you trust in this OC in a time of need?' },
  { id: 'combat', text: 'Would you follow this OC into combat?' },
  { id: 'team', text: 'When not in a leadership role, are they a good team player?' }
];

/** Score labels as the sheet spells them (index = points). */
export const SCALE = [
  '',
  'Far Below Expectations',
  'Below Expectations',
  'Meet Expectations',
  'Exceeds Expectations',
  'Far Exceeds Expectations'
];

/** Short labels for the phone screens (index = points). */
export const SCALE_SHORT = ['', 'far below', 'below', 'meets', 'exceeds', 'far exceeds'];

/** The ALRM categories, in sheet order of first appearance. */
export const CATEGORIES = ['Character', 'Presence', 'Intellect', 'Leads', 'Develops', 'Achieves'];

/** The 30 attributes. Index in this array is the attribute id used everywhere. */
export const ATTRIBUTES = [
  { name: 'Loyalty', cat: 'Character', def: 'Stays true to the unit and fellow Soldiers.' },
  { name: 'Duty', cat: 'Character', def: 'Does every task fully, without being told.' },
  { name: 'Respect', cat: 'Character', def: 'Treats every person with dignity.' },
  { name: 'Selfless Service', cat: 'Character', def: 'Puts the team and mission before self.' },
  { name: 'Honor', cat: 'Character', def: 'Lives the Army Values in every action.' },
  { name: 'Integrity', cat: 'Character', def: 'Honest, and does what is right, always.' },
  { name: 'Personal Courage', cat: 'Character', def: 'Faces fear and stands up for what is right.' },
  { name: 'Warrior Ethos', cat: 'Character', def: 'Mission first; never quits; leaves no one.' },
  { name: 'Empathy', cat: 'Character', def: 'Understands what others feel and need.' },
  { name: 'Discipline', cat: 'Character', def: 'Does what is right, even when no one watches.' },
  { name: 'Military Bearing', cat: 'Presence', def: 'Looks, speaks and acts like a leader.' },
  { name: 'Fitness', cat: 'Presence', def: 'Fit enough to lead under stress.' },
  { name: 'Confidence', cat: 'Presence', def: 'Projects calm certainty that we will succeed.' },
  { name: 'Resilience', cat: 'Presence', def: 'Bounces back fast from setbacks and stress.' },
  { name: 'Mental Agility', cat: 'Intellect', def: 'Adapts quickly when the plan changes.' },
  { name: 'Interpersonal Tact', cat: 'Intellect', def: 'Reads people and handles them well.' },
  { name: 'Sound Judgment', cat: 'Intellect', def: 'Reads situations well and makes sound calls.' },
  { name: 'Innovation', cat: 'Intellect', def: 'Brings new ideas that solve real problems.' },
  { name: 'Expertise', cat: 'Intellect', def: 'Knows the job and the tasks at hand.' },
  { name: 'Leads Others', cat: 'Leads', def: 'Gives clear direction and gets people moving.' },
  { name: 'Extends Influence Beyond Chain of Command', cat: 'Leads', def: 'Builds cooperation beyond own team.' },
  { name: 'Leads by Example', cat: 'Leads', def: 'Sets the standard by doing it first.' },
  { name: 'Builds Trust', cat: 'Leads', def: 'Keeps promises and treats people fairly.' },
  { name: 'Communicates', cat: 'Leads', def: 'Speaks and listens so the team understands.' },
  { name: 'Creates Positive Environment', cat: 'Develops', def: 'Builds a team people want to be part of.' },
  { name: 'Develops Others', cat: 'Develops', def: 'Helps peers get better at their jobs.' },
  { name: 'Prepares Self', cat: 'Develops', def: 'Comes ready: studied, rested, equipped.' },
  { name: 'Stewards Profession', cat: 'Develops', def: "Cares for people, gear and the Army's trust." },
  { name: 'Gets Results', cat: 'Achieves', def: 'Gets the mission done, to standard.' },
  { name: 'Humility', cat: 'Character', def: 'Open to feedback; no ego.' }
];

export const ATTRIBUTE_COUNT = ATTRIBUTES.length;
