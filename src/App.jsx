import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

const BLUE = "#007AFF";
const BLACK = "#000000";

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

function addDays(dateStr, days) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}
function addMonths(dateStr, months) {
  const d = new Date(dateStr);
  d.setMonth(d.getMonth() + Number(months));
  return d.toISOString().slice(0, 10);
}
function daysBetween(a, b) {
  return Math.round((new Date(a) - new Date(b)) / 86400000);
}
function today() {
  return new Date().toISOString().slice(0, 10);
}
function computeStatus(c) {
  if (c.manualStatus === "Frozen" || c.manualStatus === "Cancelled") return c.manualStatus;
  return daysBetween(c.endDate, today()) < 0 ? "Expired" : "Active";
}
function bmiOf(heightCm, weightKg) {
  if (!heightCm || !weightKg) return "";
  const m = heightCm / 100;
  return (weightKg / (m * m)).toFixed(1);
}
function draftMessage(c) {
  return `Hi ${c.name}, this is a reminder from TrainWithSK that your membership ends on ${c.endDate}. Please renew soon to keep your access uninterrupted. Thank you!`;
}

// Convert between the app's camelCase objects and the database's snake_case columns
function toRow(c, category) {
  return {
    id: c.id,
    category,
    photo: c.photo || null,
    name: c.name,
    phone: c.phone,
    start_date: c.startDate,
    duration_months: Number(c.durationMonths || 1),
    bonus_days: Number(c.bonusDays || 0),
    end_date: c.endDate,
    payment_mode: c.paymentMode,
    amount_paid: c.amountPaid ? Number(c.amountPaid) : null,
    client_type: c.type,
    manual_status: c.manualStatus,
    remarks: c.remarks || null,
    emergency_contact: c.emergencyContact || null,
    referred_by: c.referredBy || null,
    age: c.age ? Number(c.age) : null,
    gender: c.gender || null,
    height: c.height ? Number(c.height) : null,
    weight: c.weight ? Number(c.weight) : null,
    fitness_goal: c.fitnessGoal || null,
    medical: c.medical || null,
    payments: c.payments || [],
    progress: c.progress || [],
  };
}
function fromRow(r) {
  return {
    id: r.id,
    photo: r.photo || "",
    name: r.name,
    phone: r.phone,
    startDate: r.start_date,
    durationMonths: r.duration_months,
    bonusDays: r.bonus_days,
    endDate: r.end_date,
    paymentMode: r.payment_mode,
    amountPaid: r.amount_paid,
    type: r.client_type,
    manualStatus: r.manual_status,
    remarks: r.remarks || "",
    emergencyContact: r.emergency_contact || "",
    referredBy: r.referred_by || "",
    age: r.age,
    gender: r.gender || "Male",
    height: r.height,
    weight: r.weight,
    fitnessGoal: r.fitness_goal || "General fitness",
    medical: r.medical || "",
    payments: r.payments || [],
    progress: r.progress || [],
  };
}

const emptyClient = {
  id: "",
  photo: "",
  name: "",
  phone: "",
  startDate: today(),
  durationMonths: 1,
  bonusDays: 0,
  endDate: "",
  paymentMode: "Cash",
  amountPaid: "",
  type: "New",
  manualStatus: "Active",
  remarks: "",
  emergencyContact: "",
  referredBy: "",
  payments: [],
  age: "",
  gender: "Male",
  height: "",
  weight: "",
  fitnessGoal: "General fitness",
  medical: "",
  progress: [],
};

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label style={{ display: "block", fontSize: 12, color: "#666", marginBottom: 4 }}>{label}</label>
      {children}
    </div>
  );
}
const inputStyle = { width: "100%", padding: "8px 10px", border: "1px solid #ddd", borderRadius: 6, fontSize: 14, boxSizing: "border-box" };
const btn = (bg, color) => ({ background: bg, color: color || "#fff", border: "none", borderRadius: 6, padding: "8px 14px", fontSize: 13, fontWeight: 500, cursor: "pointer" });
const btnOutline = { background: "#fff", color: BLACK, border: "1px solid #ccc", borderRadius: 6, padding: "8px 14px", fontSize: 13, cursor: "pointer" };

function Modal({ title, onClose, children, width = 480 }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "40px 16px", zIndex: 50, overflowY: "auto" }}>
      <div style={{ background: "#fff", borderRadius: 10, width, maxWidth: "100%", padding: 20, maxHeight: "85vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontFamily: "Montserrat, sans-serif" }}>{title}</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 18, cursor: "pointer" }}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState("regular");
  const [clients, setClients] = useState([]);
  const [ptClients, setPtClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [errorMsg, setErrorMsg] = useState("");
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [modal, setModal] = useState(null);
  const [ledgerFor, setLedgerFor] = useState(null);
  const [progressFor, setProgressFor] = useState(null);

  async function loadAll() {
    setLoading(true);
    setErrorMsg("");
    const { data, error } = await supabase.from("clients").select("*").order("created_at", { ascending: true });
    if (error) {
      setErrorMsg("Could not load data: " + error.message);
    } else {
      setClients(data.filter((r) => r.category === "regular").map(fromRow));
      setPtClients(data.filter((r) => r.category === "pt").map(fromRow));
    }
    setLoading(false);
  }

  useEffect(() => {
    loadAll();
  }, []);

  const list = tab === "regular" ? clients : ptClients;
  const allClients = [...clients, ...ptClients];

  const filtered = useMemo(() => {
    return list.filter((c) => {
      const q = search.toLowerCase();
      const matchesSearch = !q || c.name.toLowerCase().includes(q) || c.phone.includes(q);
      const matchesStatus = statusFilter === "All" || computeStatus(c) === statusFilter;
      return matchesSearch && matchesStatus;
    });
  }, [list, search, statusFilter]);

  const stats = useMemo(() => {
    const active = allClients.filter((c) => computeStatus(c) === "Active").length;
    const expiringSoon = allClients.filter((c) => {
      const d = daysBetween(c.endDate, today());
      return d >= 0 && d <= 7;
    }).length;
    const thisMonth = today().slice(0, 7);
    const revenue = allClients.reduce((sum, c) => {
      const paid = (c.payments || []).filter((p) => p.date.slice(0, 7) === thisMonth).reduce((s, p) => s + Number(p.amount), 0);
      const base = c.startDate.slice(0, 7) === thisMonth ? Number(c.amountPaid || 0) : 0;
      return sum + paid + base;
    }, 0);
    const newC = allClients.filter((c) => c.type === "New").length;
    const renewedC = allClients.filter((c) => c.type === "Renewed").length;
    return { active, expiringSoon, revenue, newC, renewedC };
  }, [allClients]);

  const remindersDue = allClients.filter((c) => [0, 1, 2].includes(daysBetween(c.endDate, today())));

  function openAdd() {
    setModal({ mode: "add", data: { ...emptyClient, id: uid(), startDate: today() } });
  }
  function openEdit(c) {
    setModal({ mode: "edit", data: { ...c } });
  }

  async function removeClient(id) {
    if (!window.confirm("Remove this client record? This cannot be undone.")) return;
    const { error } = await supabase.from("clients").delete().eq("id", id);
    if (error) return setErrorMsg("Could not delete: " + error.message);
    await loadAll();
  }

  async function saveClient(data) {
    const endDate = addDays(addMonths(data.startDate, data.durationMonths), Number(data.bonusDays || 0));
    const toSave = { ...data, endDate };
    const category = tab;
    const row = toRow(toSave, category);

    const { error } = await supabase.from("clients").upsert(row, { onConflict: "id" });
    if (error) return setErrorMsg("Could not save: " + error.message);

    if (modal.mode === "add" && data.referredBy && data.referredBy.trim()) {
      const refName = data.referredBy.trim();
      const { data: matches } = await supabase.from("clients").select("*").ilike("name", refName);
      if (matches && matches.length > 0) {
        const ref = matches[0];
        const newBonus = Number(ref.bonus_days || 0) + 5;
        const newEnd = addDays(addMonths(ref.start_date, ref.duration_months), newBonus);
        await supabase.from("clients").update({ bonus_days: newBonus, end_date: newEnd }).eq("id", ref.id);
      }
    }
    setModal(null);
    await loadAll();
  }

  async function addPayment(client, amount, mode, date) {
    const payment = { id: uid(), date, amount: Number(amount), mode };
    const updated = [...(client.payments || []), payment];
    const { error } = await supabase.from("clients").update({ payments: updated }).eq("id", client.id);
    if (error) return setErrorMsg("Could not add payment: " + error.message);
    await loadAll();
    setLedgerFor((prev) => (prev ? { ...prev, payments: updated } : prev));
  }
  async function addProgress(client, weight, date) {
    const bmi = bmiOf(client.height, weight);
    const entry = { id: uid(), date, weight: Number(weight), bmi };
    const updated = [...(client.progress || []), entry];
    const { error } = await supabase.from("clients").update({ progress: updated }).eq("id", client.id);
    if (error) return setErrorMsg("Could not add entry: " + error.message);
    await loadAll();
    setProgressFor((prev) => (prev ? { ...prev, progress: updated } : prev));
  }

  function exportExcel() {
    const rows = allClients.map((c) => ({
      Name: c.name, Phone: c.phone, Start: c.startDate, End: c.endDate,
      Duration_Months: c.durationMonths, Bonus_Days: c.bonusDays,
      Payment_Mode: c.paymentMode, Amount_Paid: c.amountPaid, New_Renewed: c.type,
      Status: computeStatus(c), Referred_By: c.referredBy, Emergency_Contact: c.emergencyContact, Remarks: c.remarks,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Clients");
    XLSX.writeFile(wb, "trainwithsk_clients.xlsx");
  }

  function copyMessage(c) {
    const msg = draftMessage(c);
    navigator.clipboard?.writeText(msg);
    alert("Message copied — paste it into your messaging app for " + c.phone);
  }

  return (
    <div style={{ fontFamily: "Montserrat, sans-serif", color: BLACK, padding: 16, maxWidth: 1100, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div style={{ fontFamily: "Audiowide, sans-serif", fontSize: 22 }}>
          TrainWith<span style={{ color: BLUE }}>_SK</span>
        </div>
        <div style={{ fontSize: 13, color: "#666" }}>Admin dashboard</div>
      </div>

      {errorMsg && (
        <div style={{ background: "#fdecea", color: "#b42318", borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 13 }}>
          {errorMsg}
        </div>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 20, borderBottom: "1px solid #eee" }}>
        {["regular", "pt"].map((t) => (
          <div key={t} onClick={() => setTab(t)}
            style={{ padding: "8px 16px", cursor: "pointer", fontSize: 14, fontWeight: 500, borderBottom: tab === t ? `2px solid ${BLUE}` : "2px solid transparent", color: tab === t ? BLACK : "#888" }}>
            {t === "regular" ? "Regular clients" : "PT clients"}
          </div>
        ))}
      </div>

      {loading ? (
        <p style={{ color: "#999" }}>Loading your data…</p>
      ) : (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 18 }}>
            {[
              ["Active members", stats.active],
              ["Expiring in 7 days", stats.expiringSoon],
              ["Revenue this month", "₹" + stats.revenue.toLocaleString()],
              ["New / Renewed", `${stats.newC} / ${stats.renewedC}`],
            ].map(([label, val]) => (
              <div key={label} style={{ background: "#f7f7f8", borderRadius: 8, padding: "12px 14px" }}>
                <div style={{ fontSize: 12, color: "#777" }}>{label}</div>
                <div style={{ fontSize: 20, fontWeight: 600, marginTop: 4 }}>{val}</div>
              </div>
            ))}
          </div>

          {remindersDue.length > 0 && (
            <div style={{ background: "#fff6e5", borderRadius: 8, padding: "10px 14px", marginBottom: 18 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Membership reminders due</div>
              {remindersDue.map((c) => (
                <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "4px 0" }}>
                  <span>{c.name} — expires {c.endDate} ({c.phone})</span>
                  <button style={btnOutline} onClick={() => copyMessage(c)}>Copy message</button>
                </div>
              ))}
            </div>
          )}

          <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
            <input style={{ ...inputStyle, flex: 1, minWidth: 180 }} placeholder="Search by name or phone" value={search} onChange={(e) => setSearch(e.target.value)} />
            <select style={{ ...inputStyle, width: 140 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
              {["All", "Active", "Expired", "Frozen", "Cancelled"].map((s) => <option key={s}>{s}</option>)}
            </select>
            <button style={btnOutline} onClick={exportExcel}>Export to Excel</button>
            <button style={btnOutline} onClick={() => window.print()}>Export to PDF (print)</button>
            <button style={btn(BLUE)} onClick={openAdd}>+ Add {tab === "regular" ? "client" : "PT client"}</button>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: "1px solid #eee", textAlign: "left", color: "#777" }}>
                  <th style={{ padding: 8 }}>Client</th>
                  <th style={{ padding: 8 }}>Phone</th>
                  <th style={{ padding: 8 }}>Start</th>
                  <th style={{ padding: 8 }}>End</th>
                  <th style={{ padding: 8 }}>Payment</th>
                  <th style={{ padding: 8 }}>Type</th>
                  {tab === "pt" && <th style={{ padding: 8 }}>BMI</th>}
                  <th style={{ padding: 8 }}>Status</th>
                  <th style={{ padding: 8 }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((c) => {
                  const status = computeStatus(c);
                  const statusColor = { Active: "#1a7f37", Expired: "#b42318", Frozen: "#946800", Cancelled: "#666" }[status];
                  return (
                    <tr key={c.id} style={{ borderBottom: "1px solid #f2f2f2" }}>
                      <td style={{ padding: 8, display: "flex", alignItems: "center", gap: 8 }}>
                        {c.photo ? (
                          <img src={c.photo} alt="" style={{ width: 26, height: 26, borderRadius: "50%", objectFit: "cover" }} />
                        ) : (
                          <div style={{ width: 26, height: 26, borderRadius: "50%", background: "#e6f1fb", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: BLUE, fontWeight: 600 }}>
                            {c.name.slice(0, 2).toUpperCase()}
                          </div>
                        )}
                        {c.name}
                      </td>
                      <td style={{ padding: 8 }}>{c.phone}</td>
                      <td style={{ padding: 8 }}>{c.startDate}</td>
                      <td style={{ padding: 8 }}>{c.endDate} {c.bonusDays > 0 && <span style={{ color: BLUE, fontSize: 11 }}> (+{c.bonusDays}d)</span>}</td>
                      <td style={{ padding: 8 }}>{c.paymentMode} · ₹{c.amountPaid}</td>
                      <td style={{ padding: 8 }}>{c.type}</td>
                      {tab === "pt" && <td style={{ padding: 8 }}>{bmiOf(c.height, c.weight)}</td>}
                      <td style={{ padding: 8, color: statusColor, fontWeight: 500 }}>{status}</td>
                      <td style={{ padding: 8, display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <button style={btnOutline} onClick={() => openEdit(c)}>Edit</button>
                        <button style={btnOutline} onClick={() => setLedgerFor(c)}>Ledger</button>
                        {tab === "pt" && <button style={btnOutline} onClick={() => setProgressFor(c)}>Progress</button>}
                        <button style={{ ...btnOutline, color: "#b42318" }} onClick={() => removeClient(c.id)}>Delete</button>
                      </td>
                    </tr>
                  );
                })}
                {filtered.length === 0 && (
                  <tr><td colSpan={9} style={{ padding: 20, textAlign: "center", color: "#999" }}>No clients yet. Add one to get started.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}

      {modal && <ClientModal isPT={tab === "pt"} mode={modal.mode} data={modal.data} onClose={() => setModal(null)} onSave={saveClient} />}
      {ledgerFor && <LedgerModal client={ledgerFor} onClose={() => setLedgerFor(null)} onAdd={addPayment} />}
      {progressFor && <ProgressModal client={progressFor} onClose={() => setProgressFor(null)} onAdd={addProgress} />}
    </div>
  );
}

function ClientModal({ isPT, mode, data, onClose, onSave }) {
  const [form, setForm] = useState(data);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  return (
    <Modal title={(mode === "add" ? "Add " : "Edit ") + (isPT ? "PT client" : "client")} onClose={onClose} width={560}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <div style={{ gridColumn: "1 / -1", display: "flex", alignItems: "center", gap: 12 }}>
          {form.photo ? (
            <img src={form.photo} alt="" style={{ width: 56, height: 56, borderRadius: "50%", objectFit: "cover" }} />
          ) : (
            <div style={{ width: 56, height: 56, borderRadius: "50%", background: "#e6f1fb", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, color: BLUE, fontWeight: 600 }}>
              {(form.name || "?").slice(0, 2).toUpperCase()}
            </div>
          )}
          <div>
            <label style={{ ...btnOutline, display: "inline-block" }}>
              Upload photo
              <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => {
                const file = e.target.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = () => set("photo", reader.result);
                reader.readAsDataURL(file);
              }} />
            </label>
            {form.photo && <button style={{ ...btnOutline, marginLeft: 8 }} onClick={() => set("photo", "")}>Remove</button>}
          </div>
        </div>
        <Field label="Client name"><input style={inputStyle} value={form.name} onChange={(e) => set("name", e.target.value)} /></Field>
        <Field label="Phone number"><input style={inputStyle} value={form.phone} onChange={(e) => set("phone", e.target.value)} /></Field>
        <Field label="Membership start date"><input type="date" style={inputStyle} value={form.startDate} onChange={(e) => set("startDate", e.target.value)} /></Field>
        <Field label="Duration (months)"><input type="number" min="1" style={inputStyle} value={form.durationMonths} onChange={(e) => set("durationMonths", e.target.value)} /></Field>
        <Field label="Mode of payment">
          <select style={inputStyle} value={form.paymentMode} onChange={(e) => set("paymentMode", e.target.value)}>
            {["Cash", "UPI", "Card", "Bank transfer"].map((m) => <option key={m}>{m}</option>)}
          </select>
        </Field>
        <Field label="Amount paid"><input type="number" style={inputStyle} value={form.amountPaid} onChange={(e) => set("amountPaid", e.target.value)} /></Field>
        <Field label="New or renewed">
          <select style={inputStyle} value={form.type} onChange={(e) => set("type", e.target.value)}>
            {["New", "Renewed"].map((m) => <option key={m}>{m}</option>)}
          </select>
        </Field>
        <Field label="Status">
          <select style={inputStyle} value={form.manualStatus} onChange={(e) => set("manualStatus", e.target.value)}>
            {["Active", "Frozen", "Cancelled"].map((m) => <option key={m}>{m}</option>)}
          </select>
        </Field>
        <Field label="Emergency contact"><input style={inputStyle} value={form.emergencyContact} onChange={(e) => set("emergencyContact", e.target.value)} /></Field>
        <Field label="Referred by (existing client name)"><input style={inputStyle} value={form.referredBy} onChange={(e) => set("referredBy", e.target.value)} /></Field>
        {isPT && (
          <>
            <Field label="Age"><input type="number" style={inputStyle} value={form.age} onChange={(e) => set("age", e.target.value)} /></Field>
            <Field label="Gender">
              <select style={inputStyle} value={form.gender} onChange={(e) => set("gender", e.target.value)}>
                {["Male", "Female", "Other"].map((g) => <option key={g}>{g}</option>)}
              </select>
            </Field>
            <Field label="Height (cm)"><input type="number" style={inputStyle} value={form.height} onChange={(e) => set("height", e.target.value)} /></Field>
            <Field label="Weight (kg)"><input type="number" style={inputStyle} value={form.weight} onChange={(e) => set("weight", e.target.value)} /></Field>
            <Field label="BMI (auto-calculated)"><input style={{ ...inputStyle, background: "#f5f5f5" }} value={bmiOf(form.height, form.weight)} readOnly /></Field>
            <Field label="Fitness goal">
              <select style={inputStyle} value={form.fitnessGoal} onChange={(e) => set("fitnessGoal", e.target.value)}>
                {["Weight loss", "Muscle gain", "Rehab", "General fitness"].map((g) => <option key={g}>{g}</option>)}
              </select>
            </Field>
            <Field label="Medical conditions / injuries"><input style={inputStyle} value={form.medical} onChange={(e) => set("medical", e.target.value)} /></Field>
          </>
        )}
        <div style={{ gridColumn: "1 / -1" }}>
          <Field label="Remarks / notes"><textarea style={{ ...inputStyle, minHeight: 60 }} value={form.remarks} onChange={(e) => set("remarks", e.target.value)} /></Field>
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <button style={btnOutline} onClick={onClose}>Cancel</button>
        <button style={btn(BLUE)} onClick={() => onSave(form)} disabled={!form.name || !form.phone}>Save client</button>
      </div>
    </Modal>
  );
}

function LedgerModal({ client, onClose, onAdd }) {
  const [amount, setAmount] = useState("");
  const [mode, setMode] = useState("Cash");
  const [date, setDate] = useState(today());
  return (
    <Modal title={"Payment history — " + client.name} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        {(client.payments || []).length === 0 && <p style={{ color: "#999", fontSize: 13 }}>No additional payments logged yet.</p>}
        {(client.payments || []).map((p) => (
          <div key={p.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "6px 0", borderBottom: "1px solid #f2f2f2" }}>
            <span>{p.date}</span><span>{p.mode}</span><span>₹{p.amount}</span>
          </div>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
        <Field label="Date"><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="Amount"><input type="number" style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} /></Field>
        <Field label="Mode">
          <select style={inputStyle} value={mode} onChange={(e) => setMode(e.target.value)}>
            {["Cash", "UPI", "Card", "Bank transfer"].map((m) => <option key={m}>{m}</option>)}
          </select>
        </Field>
        <button style={btn(BLUE)} onClick={() => { if (amount) { onAdd(client, amount, mode, date); setAmount(""); } }}>Add</button>
      </div>
    </Modal>
  );
}

function ProgressModal({ client, onClose, onAdd }) {
  const [weight, setWeight] = useState("");
  const [date, setDate] = useState(today());
  return (
    <Modal title={"Progress tracking — " + client.name} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        {(client.progress || []).length === 0 && <p style={{ color: "#999", fontSize: 13 }}>No entries yet.</p>}
        {(client.progress || []).map((p) => (
          <div key={p.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "6px 0", borderBottom: "1px solid #f2f2f2" }}>
            <span>{p.date}</span><span>{p.weight} kg</span><span>BMI {p.bmi}</span>
          </div>
        ))}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
        <Field label="Date"><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="Weight (kg)"><input type="number" style={inputStyle} value={weight} onChange={(e) => setWeight(e.target.value)} /></Field>
        <button style={btn(BLUE)} onClick={() => { if (weight) { onAdd(client, weight, date); setWeight(""); } }}>Add entry</button>
      </div>
    </Modal>
  );
}
