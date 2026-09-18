import * as React from 'react';
import { useState, useEffect } from 'react';
import { IPermissionsManagerMkIProps } from './IPermissionsManagerMkIProps';
import { getSP, logAction } from '../pnpjsConfig';
import styles from './PermissionsManagerMkI.module.scss';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/security/list";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";



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
const GROUPE_ASSOCIES = "Associés"; // À adapter au nom exact du groupe SharePoint

const PermissionsManagerMkI: React.FC<IPermissionsManagerMkIProps> = (props) => {
  const [selectedLibrary, setSelectedLibrary] = useState<string>("");
  const [permissions, setPermissions] = useState<IPermissionEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");

  const [roleDefs, setRoleDefs] = useState<IRoleDef[]>([]);
  const [emailToAdd, setEmailToAdd] = useState<string>("");
  const [selectedRoleId, setSelectedRoleId] = useState<number | "">("");
  const [addLoading, setAddLoading] = useState<boolean>(false);
  const [addMessage, setAddMessage] = useState<string>("");
  const [removingId, setRemovingId] = useState<number | null>(null);

  const [authChecked, setAuthChecked] = useState<boolean>(false);
  const [isAuthorized, setIsAuthorized] = useState<boolean>(false);

  const sp = getSP(props.context);

  useEffect(() => {
    sp.web.currentUser.groups().then((groups: any[]) => {
      const autorise = groups.some((g) => g.Title === GROUPE_ASSOCIES);
      setIsAuthorized(autorise);
      setAuthChecked(true);
    }).catch((e) => {
      console.error("Impossible de vérifier les groupes de l'utilisateur", e);
      setIsAuthorized(false);
      setAuthChecked(true);
    });
  }, []);

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
      await loadPermissions(selectedLibrary);
    } catch (e: any) {
      console.error(e);
      setAddMessage(`Erreur lors de l'ajout : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setAddLoading(false);
    }
  };

  const handleRemovePermission = async (entry: IPermissionEntry): Promise<void> => {
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
    } catch (e: any) {
      console.error(e);
      setError(`Erreur lors de la suppression de l'accès de ${entry.nom} : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setRemovingId(null);
    }
  };

  if (!authChecked) {
    return <div className={styles.container}><p className={styles.statusText}>Vérification des autorisations...</p></div>;
  }

  if (!isAuthorized) {
    return (
      <div className={styles.container}>
        <p className={styles.unauthorized}>Vous n'avez pas accès à cet outil. Contactez la direction informatique si vous pensez qu'il s'agit d'une erreur.</p>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.masthead}>
        <h1 className={styles.title}>Permissions Manager</h1>
        <p className={styles.subtitle}>Gestion des accès aux bibliothèques par direction</p>
      </div>

      <div className={styles.tabs}>
        {DIRECTIONS.map((dir) => (
          <button
            key={dir}
            className={dir === selectedLibrary ? `${styles.tab} ${styles.tabActive}` : styles.tab}
            onClick={() => loadPermissions(dir)}
          >
            {dir}
          </button>
        ))}
      </div>

      {loading && <p className={styles.statusText}>Chargement...</p>}
      {error && <p className={styles.errorBanner}>{error}</p>}

      {!loading && !error && selectedLibrary && (
        <div>
          <h2 className={styles.libraryTitle}>{selectedLibrary}</h2>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Nom</th>
                <th>Type</th>
                <th>Niveau</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {permissions.map((p) => (
                <tr key={p.principalId}>
                  <td>{p.nom}</td>
                  <td><span className={styles.typeTag}>{p.type}</span></td>
                  <td className={p.niveau.includes("Contrôle total") ? styles.levelFullControl : undefined}>{p.niveau}</td>
                  <td>
                    <button
                      className={styles.removeButton}
                      onClick={() => handleRemovePermission(p)}
                      disabled={removingId === p.principalId}
                    >
                      {removingId === p.principalId ? "..." : "Retirer"}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className={styles.addPanel}>
            <h3 className={styles.addPanelTitle}>Ajouter un accès</h3>
            <div className={styles.addForm}>
              <input
                type="email"
                className={styles.input}
                placeholder="email@cofidestsas.com"
                value={emailToAdd}
                onChange={(e) => setEmailToAdd(e.target.value)}
              />
              <select
                className={styles.select}
                value={selectedRoleId}
                onChange={(e) => setSelectedRoleId(e.target.value ? Number(e.target.value) : "")}
              >
                <option value="">Niveau d'accès</option>
                {roleDefs.map((r) => (
                  <option key={r.Id} value={r.Id}>{r.Name}</option>
                ))}
              </select>
              <button className={styles.primaryButton} onClick={handleAddPermission} disabled={addLoading}>
                {addLoading ? "Ajout..." : "Ajouter"}
              </button>
            </div>
            {addMessage && <p className={styles.feedbackMessage}>{addMessage}</p>}
          </div>
        </div>
      )}
    </div>
  );
};

export default PermissionsManagerMkI;