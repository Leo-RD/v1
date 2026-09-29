import * as React from 'react';
import { useState, useEffect } from 'react';
import { IPermissionsManagerMkIProps } from './IPermissionsManagerMkIProps';
import { getSP, logAction } from '../pnpjsConfig';
import styles from './PermissionsManagerMkI.module.scss';
import logoCofidest from '../assets/icon-cofidest-group.png';
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

// Les onglets affichent toutes les bibliothèques de documents visibles dont le titre commence par ce préfixe
const DIRECTION_PREFIX = "Direction ";
const TEMPLATE_DOCUMENT_LIBRARY = 101;
const GROUPE_ASSOCIES = "Associés"; // À adapter au nom exact du groupe SharePoint
const CONTROLE_TOTAL = "Contrôle total";
const MESSAGE_GROUPE_PROTEGE = "Par sécurité, la suppression d'un groupe entier n'est pas autorisée ici.";
// Caractères refusés par SharePoint dans le nom (et l'URL) d'une bibliothèque
const CARACTERES_INTERDITS = /[~"#%&*:<>?/\\{|}]/;

const estGroupe = (principalType: number): boolean =>
  principalType === PRINCIPAL_TYPE_SHAREPOINT_GROUP ||
  principalType === PRINCIPAL_TYPE_SECURITY_GROUP ||
  principalType === PRINCIPAL_TYPE_DISTRIBUTION_LIST;

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
  const [successMessage, setSuccessMessage] = useState<string>("");
  const [isAddModalOpen, setIsAddModalOpen] = useState<boolean>(false);
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

  const [directions, setDirections] = useState<string[]>([]);
  const [isCreateModalOpen, setIsCreateModalOpen] = useState<boolean>(false);
  const [newDirectionName, setNewDirectionName] = useState<string>("");
  const [createLoading, setCreateLoading] = useState<boolean>(false);
  const [createMessage, setCreateMessage] = useState<string>("");

  const sp = getSP(props.context);

  const loadDirections = async (): Promise<void> => {
    try {
      const libs: any[] = await sp.web.lists
        .filter(`BaseTemplate eq ${TEMPLATE_DOCUMENT_LIBRARY} and Hidden eq false`)
        .select("Title")();
      setDirections(libs
        .map((l) => l.Title as string)
        .filter((t) => t.indexOf(DIRECTION_PREFIX) === 0)
        .sort((a, b) => a.localeCompare(b, "fr")));
    } catch (e) {
      console.error("Impossible de charger la liste des directions", e);
    }
  };

  useEffect(() => {
    loadDirections().catch(console.error);
  }, []);

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

  // Fermeture des modales avec la touche Échap (sauf pendant une opération en cours)
  useEffect(() => {
    if (!isAddModalOpen && !isCreateModalOpen) return;
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key !== "Escape") return;
      if (isAddModalOpen && !addLoading) {
        setIsAddModalOpen(false);
        setAddMessage("");
      }
      if (isCreateModalOpen && !createLoading) {
        setIsCreateModalOpen(false);
        setCreateMessage("");
      }
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [isAddModalOpen, addLoading, isCreateModalOpen, createLoading]);

  const loadPermissions = async (libraryName: string): Promise<void> => {
    setLoading(true);
    setError("");
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

  const openAddModal = (): void => {
    setTargetType("user");
    setEmailToAdd("");
    setSelectedGroupId("");
    setSelectedRoleId("");
    setAddMessage("");
    setSuccessMessage("");
    setIsAddModalOpen(true);
  };

  const closeAddModal = (): void => {
    if (addLoading) return;
    setIsAddModalOpen(false);
    setAddMessage("");
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
      await logAction(sp, "Ajout", selectedLibrary, cible, `Niveau : ${roleDefs.find(r => r.Id === selectedRoleId)?.Name ?? ""}`);
      setEmailToAdd("");
      setSelectedGroupId("");
      setSelectedRoleId("");
      setIsAddModalOpen(false);
      await loadPermissions(selectedLibrary);
      setSuccessMessage(`Accès accordé à ${cible}.`);
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
    // Garde-fou : le bouton est désactivé pour les groupes, mais on bloque aussi ici
    // (type de la ligne, ou ID présent parmi les groupes SharePoint du site)
    if (estGroupe(entry.principalType) || siteGroups.some((g) => g.Id === entry.principalId)) {
      setError(MESSAGE_GROUPE_PROTEGE);
      return;
    }

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
    setSuccessMessage("");
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

  const openCreateModal = (): void => {
    setNewDirectionName("");
    setCreateMessage("");
    setIsCreateModalOpen(true);
  };

  const closeCreateModal = (): void => {
    if (createLoading) return;
    setIsCreateModalOpen(false);
    setCreateMessage("");
  };

  const handleCreateDirection = async (): Promise<void> => {
    const saisie = newDirectionName.trim();
    if (!saisie) {
      setCreateMessage("Renseignez le nom de la nouvelle direction.");
      return;
    }
    if (CARACTERES_INTERDITS.test(saisie)) {
      setCreateMessage("Le nom ne peut pas contenir les caractères suivants : ~ \" # % & * : < > ? / \\ { | }");
      return;
    }
    // Préfixe automatique pour que la bibliothèque apparaisse dans les onglets
    const nom = saisie.toLowerCase().indexOf(DIRECTION_PREFIX.toLowerCase()) === 0
      ? DIRECTION_PREFIX + saisie.substring(DIRECTION_PREFIX.length)
      : DIRECTION_PREFIX + saisie;
    if (directions.some((d) => d.toLowerCase() === nom.toLowerCase())) {
      setCreateMessage(`La bibliothèque "${nom}" existe déjà.`);
      return;
    }

    setCreateLoading(true);
    setCreateMessage("");
    let bibliothequeCreee = false;
    try {
      const info = await sp.web.lists.add(nom, "", TEMPLATE_DOCUMENT_LIBRARY);
      bibliothequeCreee = true;
      // Étape obligatoire : sans rupture d'héritage, roleAssignments.add()/remove() renvoient une erreur 400.
      // Sans copie des permissions, SharePoint attribue le Contrôle total au créateur.
      await sp.web.lists.getById(info.Id).breakRoleInheritance(false);
      await logAction(sp, "Initialisation", nom, "", "Bibliothèque de direction créée, héritage des permissions rompu");

      setIsCreateModalOpen(false);
      await loadDirections();
      closeMembers();
      await loadPermissions(nom);
      setSuccessMessage(`La bibliothèque "${nom}" a été créée.`);
    } catch (e: any) {
      console.error(e);
      const detail = e?.message || "erreur inconnue, voir la console (F12).";
      setCreateMessage(bibliothequeCreee
        ? `La bibliothèque "${nom}" a été créée, mais la rupture de l'héritage des permissions a échoué : ${detail} Contactez l'IT avant d'y ajouter des accès.`
        : `Erreur lors de la création de la bibliothèque : ${detail}`);
      if (bibliothequeCreee) {
        await loadDirections();
      }
    } finally {
      setCreateLoading(false);
    }
  };

  const masthead = (
    <div className={styles.masthead}>
      <img className={styles.logo} src={logoCofidest} alt="Cofidest" />
      <div>
        <h1 className={styles.title}>Permissions Manager</h1>
        <p className={styles.subtitle}>Gestion des accès aux bibliothèques par direction</p>
      </div>
    </div>
  );

  if (!authChecked) {
    return (
      <div className={styles.container}>
        {masthead}
        <p className={styles.statusText}>Vérification des autorisations...</p>
      </div>
    );
  }

  if (!isAuthorized) {
    return (
      <div className={styles.container}>
        {masthead}
        <p className={styles.unauthorized}>Vous n'avez pas accès à cet outil. Contactez la direction informatique si vous pensez qu'il s'agit d'une erreur.</p>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      {masthead}

      <div className={styles.tabs}>
        {directions.map((dir) => (
          <button
            key={dir}
            className={dir === selectedLibrary ? `${styles.tab} ${styles.tabActive}` : styles.tab}
            onClick={() => { closeMembers(); setSuccessMessage(""); loadPermissions(dir).catch(console.error); }}
          >
            {dir}
          </button>
        ))}
        <button className={`${styles.tab} ${styles.newDirectionTab}`} onClick={openCreateModal}>
          <span className={styles.plusIcon} aria-hidden="true">+</span>
          Créer une nouvelle Direction
        </button>
      </div>

      {!selectedLibrary && !loading && (
        <p className={styles.emptyState}>Sélectionnez une direction pour afficher les accès à sa bibliothèque.</p>
      )}

      {loading && <p className={styles.statusText}>Chargement...</p>}
      {error && <p className={styles.errorBanner}>{error}</p>}

      {!loading && !error && selectedLibrary && (
        <div>
          <div className={styles.toolbar}>
            <h2 className={styles.libraryTitle}>{selectedLibrary}</h2>
            <button className={`${styles.primaryButton} ${styles.addAccessButton}`} onClick={openAddModal}>
              <span className={styles.plusIcon} aria-hidden="true">+</span>
              Ajouter un accès
            </button>
          </div>

          {successMessage && <p className={styles.successBanner}>{successMessage}</p>}

          <div className={styles.tableWrapper}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Nom</th>
                  <th>Type</th>
                  <th>Niveau</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {permissions.map((p) => (
                  <tr key={p.principalId}>
                    <td className={styles.nameCell}>{p.nom}</td>
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
                          disabled={estGroupe(p.principalType) || removingId === p.principalId}
                          title={estGroupe(p.principalType) ? MESSAGE_GROUPE_PROTEGE : undefined}
                        >
                          {removingId === p.principalId ? "..." : "Retirer"}
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {membersGroup && (
            <div className={styles.membersPanel}>
              <div className={styles.panelHeader}>
                <h3 className={styles.panelTitle}>Membres du groupe « {membersGroup.nom} »</h3>
                <button className={styles.secondaryButton} onClick={closeMembers}>Fermer</button>
              </div>

              {membersLoading ? (
                <p className={styles.statusText}>Chargement des membres...</p>
              ) : members.length === 0 ? (
                <p className={styles.statusText}>Ce groupe ne contient aucun membre.</p>
              ) : (
                <div className={styles.tableWrapper}>
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
                          <td className={styles.nameCell}>{m.Title}</td>
                          <td>{m.Email}</td>
                          <td><span className={styles.typeTag}>{getPrincipalTypeLabel(m.PrincipalType)}</span></td>
                          <td>
                            <div className={styles.rowActions}>
                              <button
                                className={styles.removeButton}
                                onClick={() => handleRemoveMember(m)}
                                disabled={removingMemberId === m.Id}
                              >
                                {removingMemberId === m.Id ? "..." : "Retirer"}
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              <div className={styles.inlineForm}>
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
        </div>
      )}

      {isAddModalOpen && (
        <div className={styles.modalBackdrop} onClick={closeAddModal}>
          <div
            className={styles.modal}
            role="dialog"
            aria-modal="true"
            aria-labelledby="pm-add-modal-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.modalHeader}>
              <div>
                <h3 id="pm-add-modal-title" className={styles.modalTitle}>Ajouter un accès</h3>
                <p className={styles.modalSubtitle}>{selectedLibrary}</p>
              </div>
              <button className={styles.closeButton} onClick={closeAddModal} disabled={addLoading} aria-label="Fermer">
                ×
              </button>
            </div>

            <div className={styles.modalBody}>
              <div className={styles.segmented} role="radiogroup" aria-label="Type de cible">
                <button
                  type="button"
                  role="radio"
                  aria-checked={targetType === "user"}
                  className={targetType === "user" ? `${styles.segment} ${styles.segmentActive}` : styles.segment}
                  onClick={() => handleTargetTypeChange("user")}
                >
                  Utilisateur
                </button>
                <button
                  type="button"
                  role="radio"
                  aria-checked={targetType === "group"}
                  className={targetType === "group" ? `${styles.segment} ${styles.segmentActive}` : styles.segment}
                  onClick={() => handleTargetTypeChange("group")}
                >
                  Groupe SharePoint
                </button>
              </div>

              <label className={styles.field}>
                <span className={styles.fieldLabel}>{targetType === "user" ? "Email de l'utilisateur" : "Groupe SharePoint"}</span>
                {targetType === "user" ? (
                  <input
                    type="email"
                    className={styles.input}
                    placeholder="email@cofidestsas.com"
                    value={emailToAdd}
                    onChange={(e) => setEmailToAdd(e.target.value)}
                    autoFocus
                  />
                ) : (
                  <select
                    className={styles.select}
                    value={selectedGroupId}
                    onChange={(e) => setSelectedGroupId(e.target.value ? Number(e.target.value) : "")}
                  >
                    <option value="">Choisir un groupe</option>
                    {siteGroups.map((g) => (
                      <option key={g.Id} value={g.Id}>{g.Title}</option>
                    ))}
                  </select>
                )}
              </label>

              <label className={styles.field}>
                <span className={styles.fieldLabel}>Niveau d'accès</span>
                <select
                  className={styles.select}
                  value={selectedRoleId}
                  onChange={(e) => setSelectedRoleId(e.target.value ? Number(e.target.value) : "")}
                >
                  <option value="">Choisir un niveau</option>
                  {roleDefs.map((r) => (
                    <option key={r.Id} value={r.Id}>{r.Name}</option>
                  ))}
                </select>
              </label>

              {addMessage && <p className={styles.modalError}>{addMessage}</p>}
            </div>

            <div className={styles.modalFooter}>
              <button className={styles.secondaryButton} onClick={closeAddModal} disabled={addLoading}>
                Annuler
              </button>
              <button className={styles.primaryButton} onClick={handleAddPermission} disabled={addLoading}>
                {addLoading ? "Ajout..." : "Ajouter l'accès"}
              </button>
            </div>
          </div>
        </div>
      )}

      {isCreateModalOpen && (
        <div className={styles.modalBackdrop} onClick={closeCreateModal}>
          <div
            className={styles.modal}
            role="dialog"
            aria-modal="true"
            aria-labelledby="pm-create-modal-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.modalHeader}>
              <div>
                <h3 id="pm-create-modal-title" className={styles.modalTitle}>Créer une nouvelle Direction</h3>
                <p className={styles.modalSubtitle}>Nouvelle bibliothèque de documents, aux permissions indépendantes du site</p>
              </div>
              <button className={styles.closeButton} onClick={closeCreateModal} disabled={createLoading} aria-label="Fermer">
                ×
              </button>
            </div>

            <div className={styles.modalBody}>
              <label className={styles.field}>
                <span className={styles.fieldLabel}>Nom de la direction</span>
                <input
                  type="text"
                  className={styles.input}
                  placeholder="Direction Juridique"
                  value={newDirectionName}
                  onChange={(e) => setNewDirectionName(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter" && !createLoading) handleCreateDirection().catch(console.error); }}
                  autoFocus
                />
              </label>

              {createMessage && <p className={styles.modalError}>{createMessage}</p>}
            </div>

            <div className={styles.modalFooter}>
              <button className={styles.secondaryButton} onClick={closeCreateModal} disabled={createLoading}>
                Annuler
              </button>
              <button className={styles.primaryButton} onClick={handleCreateDirection} disabled={createLoading}>
                {createLoading ? "Création..." : "Créer la bibliothèque"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PermissionsManagerMkI;
