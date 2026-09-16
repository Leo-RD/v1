import * as React from 'react';
import { useState, useEffect } from 'react';
import { IPermissionsManagerMkIProps } from './IPermissionsManagerMkIProps';
import { getSP, logAction } from '../pnpjsConfig';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/security/list";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";

interface IPermissionEntry {
  principalId: number;
  nom: string;
  type: string;
  niveau: string;
  roleDefIds: number[];
}

interface IRoleDef {
  Id: number;
  Name: string;
}

const DIRECTIONS = ["Direction Informatique", "Direction Financiere", "Direction RH"];

const PermissionsManagerMkI: React.FC<IPermissionsManagerMkIProps> = (props) => {
  const [selectedLibrary, setSelectedLibrary] = useState<string>("");
  const [permissions, setPermissions] = useState<IPermissionEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");

  // --- Nouveau : état pour le formulaire d'ajout ---
  const [roleDefs, setRoleDefs] = useState<IRoleDef[]>([]);
  const [emailToAdd, setEmailToAdd] = useState<string>("");
  const [selectedRoleId, setSelectedRoleId] = useState<number | "">("");
  const [addLoading, setAddLoading] = useState<boolean>(false);
  const [addMessage, setAddMessage] = useState<string>("");
  const [removingId, setRemovingId] = useState<number | null>(null);

  const sp = getSP(props.context);

  // Récupère une seule fois les niveaux d'accès disponibles sur le site
  useEffect(() => {
    sp.web.roleDefinitions().then((defs: any[]) => {
      setRoleDefs(defs.map((d) => ({ Id: d.Id, Name: d.Name })));
    }).catch((e) => console.error("Impossible de charger les niveaux d'accès", e));
  }, []);

  const loadPermissions = async (libraryName: string): Promise<void> => {
    setLoading(true);
    setError("");
    setAddMessage("");
    setSelectedLibrary(libraryName);
    try {
      const roleAssignments: any[] = await sp.web.lists
        .getByTitle(libraryName)
        .roleAssignments.expand("Member", "RoleDefinitionBindings")();

      const mapped: IPermissionEntry[] = roleAssignments.map((ra) => ({
        principalId: ra.PrincipalId,
        nom: ra.Member?.Title ?? "Inconnu",
        type: ra.Member?.PrincipalType === 8 ? "Groupe" : "Utilisateur",
        niveau: (ra.RoleDefinitionBindings || []).map((r: any) => r.Name).join(", "),
        roleDefIds: (ra.RoleDefinitionBindings || []).map((r: any) => r.Id)
      }));
        setPermissions(mapped);
        await logAction(sp, "Consultation", libraryName);
    } catch (e) {
      console.error(e);
      setError("Impossible de récupérer les permissions. Vérifiez le nom exact de la bibliothèque et vos droits d'accès.");
    } finally {
      setLoading(false);
    }
  };

  // --- Nouveau : ajout d'une permission ---
  const handleAddPermission = async (): Promise<void> => {
    if (!emailToAdd || selectedRoleId === "") {
      setAddMessage("Renseignez un email et un niveau d'accès.");
      return;
    }
    setAddLoading(true);
    setAddMessage("");
    try {
      const user = await sp.web.ensureUser(emailToAdd);
      await sp.web.lists.getByTitle(selectedLibrary).roleAssignments.add(user.Id, selectedRoleId as number);
      setAddMessage(`Accès accordé à ${emailToAdd}.`);
      await logAction(sp, "Ajout", selectedLibrary, emailToAdd, `Niveau : ${roleDefs.find(r => r.Id === selectedRoleId)?.Name ?? ""}`);
      setEmailToAdd("");
      setSelectedRoleId("");
      await loadPermissions(selectedLibrary); // rafraîchit le tableau
    } catch (e) {
      console.error(e);
      setAddMessage("Erreur lors de l'ajout. Vérifiez que l'adresse email correspond bien à un compte du tenant.");
    } finally {
      setAddLoading(false);
    }
  };

  // --- Nouveau : suppression d'une permission ---
  const handleRemovePermission = async (entry: IPermissionEntry): Promise<void> => {
    // Garde-fou : ne pas retirer la dernière personne en "Contrôle total"
    const estSeulControleTotal =
      entry.niveau.includes("Contrôle total") &&
      permissions.filter((p) => p.niveau.includes("Contrôle total")).length === 1;

    if (estSeulControleTotal) {
      setError(`Impossible de retirer ${entry.nom} : c'est la seule personne en "Contrôle total" sur cette bibliothèque. Ajoutez d'abord quelqu'un d'autre en Contrôle total avant de le retirer.`);
      return;
    }

    const confirme = window.confirm(`Retirer tous les accès de "${entry.nom}" sur ${selectedLibrary} ?`);
    if (!confirme) return;

    setRemovingId(entry.principalId);
    setError("");
    try {
      const list = sp.web.lists.getByTitle(selectedLibrary);
            for (const roleDefId of entry.roleDefIds) {
        await list.roleAssignments.remove(entry.principalId, roleDefId);
      }
      await logAction(sp, "Suppression", selectedLibrary, entry.nom, `Niveaux retirés : ${entry.niveau}`);
      await loadPermissions(selectedLibrary);
    } catch (e) {
      console.error(e);
      setError(`Erreur lors de la suppression de l'accès de ${entry.nom}.`);
    } finally {
      setRemovingId(null);
    }
  };

  return (
    <div>
      <h2>Permissions Manager</h2>
      <div>
        {DIRECTIONS.map((dir) => (
          <button key={dir} onClick={() => loadPermissions(dir)} style={{ marginRight: 8 }}>
            {dir}
          </button>
        ))}
      </div>

      {loading && <p>Chargement...</p>}
      {error && <p style={{ color: "red" }}>{error}</p>}

      {!loading && !error && selectedLibrary && (
        <div>
          <h3>{selectedLibrary}</h3>
          <table>
            <thead>
              <tr>
                <th>Nom</th>
                <th>Type</th>
                <th>Niveau</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {permissions.map((p) => (
                <tr key={p.principalId}>
                  <td>{p.nom}</td>
                  <td>{p.type}</td>
                  <td>{p.niveau}</td>
                  <td>
                    <button onClick={() => handleRemovePermission(p)} disabled={removingId === p.principalId}>
                      {removingId === p.principalId ? "..." : "Retirer"}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* --- Formulaire d'ajout --- */}
          <div style={{ marginTop: 16, paddingTop: 16, borderTop: "1px solid #ccc" }}>
            <h4>Ajouter un accès</h4>
            <input
              type="email"
              placeholder="email@cofidestsas.com"
              value={emailToAdd}
              onChange={(e) => setEmailToAdd(e.target.value)}
              style={{ marginRight: 8 }}
            />
            <select
              value={selectedRoleId}
              onChange={(e) => setSelectedRoleId(e.target.value ? Number(e.target.value) : "")}
              style={{ marginRight: 8 }}
            >
              <option value="">-- Niveau d'accès --</option>
              {roleDefs.map((r) => (
                <option key={r.Id} value={r.Id}>{r.Name}</option>
              ))}
            </select>
            <button onClick={handleAddPermission} disabled={addLoading}>
              {addLoading ? "Ajout..." : "Ajouter"}
            </button>
            {addMessage && <p>{addMessage}</p>}
          </div>
        </div>
      )}
    </div>
  );
};

export default PermissionsManagerMkI; 