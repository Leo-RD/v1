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



// Valeurs SharePoint de PrincipalType (SP.Utilities.PrincipalType)
const PRINCIPAL_TYPE_USER = 1;
const PRINCIPAL_TYPE_DISTRIBUTION_LIST = 2;
const PRINCIPAL_TYPE_SECURITY_GROUP = 4;
const PRINCIPAL_TYPE_SHAREPOINT_GROUP = 8;

interface IPermissionEntry {
  principalId: number;
  principalType: number;
  nom: string;
  type: string;
  niveau: string;
  roleDefIds: number[];
}

interface IRoleDef {
  Id: number;
  Name: string;
}

interface ISiteGroup {
  Id: number;
  Title: string;
}

interface IGroupMember {
  Id: number;
  Title: string;
  Email: string;
  PrincipalType: number;
}

type TargetType = "user" | "group";

const DIRECTIONS = ["Direction Informatique", "Direction Financiere", "Direction RH"];
const GROUPE_ASSOCIES = "Associés"; // À adapter au nom exact du groupe SharePoint
const CONTROLE_TOTAL = "Contrôle total";

const getPrincipalTypeLabel = (principalType: number): string => {
  switch (principalType) {
    case PRINCIPAL_TYPE_SHAREPOINT_GROUP: return "Groupe SharePoint";
    case PRINCIPAL_TYPE_SECURITY_GROUP: return "Groupe de sécurité";
    case PRINCIPAL_TYPE_DISTRIBUTION_LIST: return "Liste de distribution";
    case PRINCIPAL_TYPE_USER: return "Utilisateur";
    default: return "Inconnu";
  }
};

const PermissionsManagerMkI: React.FC<IPermissionsManagerMkIProps> = (props) => {
  const [selectedLibrary, setSelectedLibrary] = useState<string>("");
  const [permissions, setPermissions] = useState<IPermissionEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");

  const [roleDefs, setRoleDefs] = useState<IRoleDef[]>([]);
  const [siteGroups, setSiteGroups] = useState<ISiteGroup[]>([]);
  const [targetType, setTargetType] = useState<TargetType>("user");
  const [emailToAdd, setEmailToAdd] = useState<string>("");
  const [selectedGroupId, setSelectedGroupId] = useState<number | "">("");
  const [selectedRoleId, setSelectedRoleId] = useState<number | "">("");
  const [addLoading, setAddLoading] = useState<boolean>(false);
  const [addMessage, setAddMessage] = useState<string>("");
  const [removingId, setRemovingId] = useState<number | null>(null);

  const [membersGroup, setMembersGroup] = useState<IPermissionEntry | null>(null);
  const [members, setMembers] = useState<IGroupMember[]>([]);
  const [membersLoading, setMembersLoading] = useState<boolean>(false);
  const [memberEmailToAdd, setMemberEmailToAdd] = useState<string>("");
  const [memberAddLoading, setMemberAddLoading] = useState<boolean>(false);
  const [memberMessage, setMemberMessage] = useState<string>("");
  const [removingMemberId, setRemovingMemberId] = useState<number | null>(null);

  const [authChecked, setAuthChecked] = useState<boolean>(false);
  const [isAuthorized, setIsAuthorized] = useState<boolean>(false);
  const [currentUserId, setCurrentUserId] = useState<number | null>(null);

  const sp = getSP(props.context);

  useEffect(() => {
    Promise.all([
      sp.web.currentUser.groups(),
      sp.web.currentUser.select("Id")()
    ]).then(([groups, me]: [any[], any]) => {
      const autorise = groups.some((g) => g.Title === GROUPE_ASSOCIES);
      setCurrentUserId(me.Id);
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

  useEffect(() => {
    sp.web.siteGroups.select("Id", "Title")().then((groups: any[]) => {
      const mapped: ISiteGroup[] = groups
        .map((g) => ({ Id: g.Id, Title: g.Title }))
        .sort((a, b) => a.Title.localeCompare(b.Title, "fr"));
      setSiteGroups(mapped);
    }).catch((e) => console.error("Impossible de charger les groupes SharePoint du site", e));
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

      const mapped: IPermissionEntry[] = roleAssignments.map((ra) => {
        const principalType: number = ra.Member?.PrincipalType ?? 0;
        return {
          principalId: ra.PrincipalId,
          principalType,
          nom: ra.Member?.Title ?? "Inconnu",
          type: getPrincipalTypeLabel(principalType),
          niveau: (ra.RoleDefinitionBindings || []).map((r: any) => r.Name).join(", "),
          roleDefIds: (ra.RoleDefinitionBindings || []).map((r: any) => r.Id)
        };
      });
      setPermissions(mapped);
      await logAction(sp, "Consultation", libraryName);
    } catch (e) {
      console.error(e);
      setError("Impossible de récupérer les permissions. Vérifiez le nom exact de la bibliothèque et vos droits d'accès.");
    } finally {
      setLoading(false);
    }
  };

  const handleTargetTypeChange = (type: TargetType): void => {
    setTargetType(type);
    setAddMessage("");
  };

  const handleAddPermission = async (): Promise<void> => {
    if (selectedRoleId === "") {
      setAddMessage("Choisissez un niveau d'accès.");
      return;
    }
    if (targetType === "user" && !emailToAdd) {
      setAddMessage("Renseignez un email et un niveau d'accès.");
      return;
    }
    const group = targetType === "group" ? siteGroups.find((g) => g.Id === selectedGroupId) : undefined;
    if (targetType === "group" && !group) {
      setAddMessage("Choisissez un groupe SharePoint et un niveau d'accès.");
      return;
    }

    setAddLoading(true);
    setAddMessage("");
    try {
      let principalId: number;
      let cible: string;
      if (group) {
        principalId = group.Id;
        cible = group.Title;
      } else {
        const user = await sp.web.ensureUser(emailToAdd);
        principalId = user.Id;
        cible = emailToAdd;
      }

      await sp.web.lists.getByTitle(selectedLibrary).roleAssignments.add(principalId, selectedRoleId as number);
      setAddMessage(`Accès accordé à ${cible}.`);
      await logAction(sp, "Ajout", selectedLibrary, cible, `Niveau : ${roleDefs.find(r => r.Id === selectedRoleId)?.Name ?? ""}`);
      setEmailToAdd("");
      setSelectedGroupId("");
      setSelectedRoleId("");
      await loadPermissions(selectedLibrary);
    } catch (e: any) {
      console.error(e);
      setAddMessage(`Erreur lors de l'ajout : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setAddLoading(false);
    }
  };

  const closeMembers = (): void => {
    setMembersGroup(null);
    setMembers([]);
    setMemberEmailToAdd("");
    setMemberMessage("");
  };

  // Le garde-fou compte tous les titulaires (utilisateurs comme groupes) du Contrôle total
  const estSeulTitulaireControleTotal = (entry: IPermissionEntry): boolean =>
    entry.niveau.includes(CONTROLE_TOTAL) &&
    permissions.filter((p) => p.niveau.includes(CONTROLE_TOTAL)).length === 1;

  const handleRemovePermission = async (entry: IPermissionEntry): Promise<void> => {
    const estSeulControleTotal = estSeulTitulaireControleTotal(entry);

    const libelleCible = entry.principalType === PRINCIPAL_TYPE_SHAREPOINT_GROUP
      ? `le groupe "${entry.nom}"`
      : entry.nom;

    if (estSeulControleTotal) {
      setError(`Impossible de retirer ${libelleCible} : c'est le seul titulaire du "${CONTROLE_TOTAL}" sur cette bibliothèque. Ajoutez d'abord un autre utilisateur ou groupe en ${CONTROLE_TOTAL} avant de le retirer.`);
      return;
    }

    const confirme = window.confirm(`Retirer tous les accès de ${libelleCible} sur ${selectedLibrary} ?`);
    if (!confirme) return;

    setRemovingId(entry.principalId);
    setError("");
    try {
      const list = sp.web.lists.getByTitle(selectedLibrary);
      for (const roleDefId of entry.roleDefIds) {
        await list.roleAssignments.remove(entry.principalId, roleDefId);
      }
      await logAction(sp, "Suppression", selectedLibrary, entry.nom, `Type : ${entry.type} - Niveaux retirés : ${entry.niveau}`);
      if (membersGroup?.principalId === entry.principalId) {
        closeMembers();
      }
      await loadPermissions(selectedLibrary);
    } catch (e: any) {
      console.error(e);
      setError(`Erreur lors de la suppression de l'accès de ${libelleCible} : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setRemovingId(null);
    }
  };

  const loadMembers = async (groupId: number): Promise<void> => {
    setMembersLoading(true);
    try {
      const users: any[] = await sp.web.siteGroups.getById(groupId)
        .users.select("Id", "Title", "Email", "PrincipalType")();
      setMembers(users
        .map((u) => ({ Id: u.Id, Title: u.Title, Email: u.Email ?? "", PrincipalType: u.PrincipalType }))
        .sort((a, b) => a.Title.localeCompare(b.Title, "fr")));
    } catch (e: any) {
      console.error(e);
      setMembers([]);
      setMemberMessage(`Impossible de charger les membres du groupe : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setMembersLoading(false);
    }
  };

  const openMembers = async (entry: IPermissionEntry): Promise<void> => {
    if (membersGroup?.principalId === entry.principalId) {
      closeMembers();
      return;
    }
    setMembersGroup(entry);
    setMemberEmailToAdd("");
    setMemberMessage("");
    await loadMembers(entry.principalId);
  };

  const handleAddMember = async (): Promise<void> => {
    if (!membersGroup) return;
    if (!memberEmailToAdd) {
      setMemberMessage("Renseignez l'email du membre à ajouter.");
      return;
    }
    setMemberAddLoading(true);
    setMemberMessage("");
    try {
      const user = await sp.web.ensureUser(memberEmailToAdd);
      await sp.web.siteGroups.getById(membersGroup.principalId).users.add(user.LoginName);
      setMemberMessage(`${memberEmailToAdd} a été ajouté au groupe "${membersGroup.nom}".`);
      await logAction(sp, "Ajout", selectedLibrary, memberEmailToAdd, `Membre ajouté au groupe SharePoint "${membersGroup.nom}"`);
      setMemberEmailToAdd("");
      await loadMembers(membersGroup.principalId);
    } catch (e: any) {
      console.error(e);
      setMemberMessage(`Erreur lors de l'ajout du membre : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setMemberAddLoading(false);
    }
  };

  const handleRemoveMember = async (member: IGroupMember): Promise<void> => {
    if (!membersGroup) return;

    if (membersGroup.nom === GROUPE_ASSOCIES && member.Id === currentUserId) {
      setMemberMessage(`Impossible de vous retirer vous-même du groupe "${GROUPE_ASSOCIES}" : vous perdriez l'accès à cet outil.`);
      return;
    }
    if (members.length === 1 && estSeulTitulaireControleTotal(membersGroup)) {
      setMemberMessage(`Impossible de retirer le dernier membre de "${membersGroup.nom}" : ce groupe est le seul titulaire du "${CONTROLE_TOTAL}" sur ${selectedLibrary}.`);
      return;
    }

    const confirme = window.confirm(`Retirer "${member.Title}" du groupe "${membersGroup.nom}" ?`);
    if (!confirme) return;

    setRemovingMemberId(member.Id);
    setMemberMessage("");
    try {
      await sp.web.siteGroups.getById(membersGroup.principalId).users.removeById(member.Id);
      await logAction(sp, "Suppression", selectedLibrary, member.Email || member.Title, `Membre retiré du groupe SharePoint "${membersGroup.nom}"`);
      await loadMembers(membersGroup.principalId);
    } catch (e: any) {
      console.error(e);
      setMemberMessage(`Erreur lors du retrait du membre : ${e?.message || "erreur inconnue, voir la console (F12)."}`);
    } finally {
      setRemovingMemberId(null);
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
            onClick={() => { closeMembers(); loadPermissions(dir).catch(console.error); }}
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
                  <td className={p.niveau.includes(CONTROLE_TOTAL) ? styles.levelFullControl : undefined}>{p.niveau}</td>
                  <td>
                    <div className={styles.rowActions}>
                      {p.principalType === PRINCIPAL_TYPE_SHAREPOINT_GROUP && (
                        <button
                          className={membersGroup?.principalId === p.principalId ? `${styles.secondaryButton} ${styles.secondaryButtonActive}` : styles.secondaryButton}
                          onClick={() => openMembers(p)}
                        >
                          Membres
                        </button>
                      )}
                      <button
                        className={styles.removeButton}
                        onClick={() => handleRemovePermission(p)}
                        disabled={removingId === p.principalId}
                      >
                        {removingId === p.principalId ? "..." : "Retirer"}
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {membersGroup && (
            <div className={`${styles.addPanel} ${styles.membersPanel}`}>
              <div className={styles.panelHeader}>
                <h3 className={styles.addPanelTitle}>Membres du groupe « {membersGroup.nom} »</h3>
                <button className={styles.secondaryButton} onClick={closeMembers}>Fermer</button>
              </div>

              {membersLoading ? (
                <p className={styles.statusText}>Chargement des membres...</p>
              ) : members.length === 0 ? (
                <p className={styles.statusText}>Ce groupe ne contient aucun membre.</p>
              ) : (
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Nom</th>
                      <th>Email</th>
                      <th>Type</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {members.map((m) => (
                      <tr key={m.Id}>
                        <td>{m.Title}</td>
                        <td>{m.Email}</td>
                        <td><span className={styles.typeTag}>{getPrincipalTypeLabel(m.PrincipalType)}</span></td>
                        <td>
                          <button
                            className={styles.removeButton}
                            onClick={() => handleRemoveMember(m)}
                            disabled={removingMemberId === m.Id}
                          >
                            {removingMemberId === m.Id ? "..." : "Retirer"}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}

              <div className={styles.addForm}>
                <input
                  type="email"
                  className={styles.input}
                  placeholder="email@cofidestsas.com"
                  value={memberEmailToAdd}
                  onChange={(e) => setMemberEmailToAdd(e.target.value)}
                />
                <button className={styles.primaryButton} onClick={handleAddMember} disabled={memberAddLoading}>
                  {memberAddLoading ? "Ajout..." : "Ajouter au groupe"}
                </button>
              </div>
              {memberMessage && <p className={styles.feedbackMessage}>{memberMessage}</p>}
            </div>
          )}

          <div className={styles.addPanel}>
            <h3 className={styles.addPanelTitle}>Ajouter un accès</h3>
            <div className={styles.targetSwitch} role="radiogroup" aria-label="Type de cible">
              <label>
                <input
                  type="radio"
                  name="targetType"
                  value="user"
                  checked={targetType === "user"}
                  onChange={() => handleTargetTypeChange("user")}
                />
                Utilisateur
              </label>
              <label>
                <input
                  type="radio"
                  name="targetType"
                  value="group"
                  checked={targetType === "group"}
                  onChange={() => handleTargetTypeChange("group")}
                />
                Groupe SharePoint
              </label>
            </div>
            <div className={styles.addForm}>
              {targetType === "user" ? (
                <input
                  type="email"
                  className={styles.input}
                  placeholder="email@cofidestsas.com"
                  value={emailToAdd}
                  onChange={(e) => setEmailToAdd(e.target.value)}
                />
              ) : (
                <select
                  className={styles.select}
                  value={selectedGroupId}
                  onChange={(e) => setSelectedGroupId(e.target.value ? Number(e.target.value) : "")}
                >
                  <option value="">Groupe SharePoint</option>
                  {siteGroups.map((g) => (
                    <option key={g.Id} value={g.Id}>{g.Title}</option>
                  ))}
                </select>
              )}
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
