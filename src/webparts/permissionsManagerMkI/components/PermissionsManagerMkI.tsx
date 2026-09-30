// Composant React principal : toute l'interface et la logique métier du gestionnaire de permissions.
// Fonctionnement :
// - Vérifie que l'utilisateur appartient au groupe SharePoint "Associés" ; sinon l'outil n'est pas affiché.
// - Liste les "directions" = bibliothèques de documents dont le titre commence par "Direction ".
// - Navigation en deux vues (V3) :
//   · Vue Accueil (aucune direction sélectionnée) : grille de cartes, une par direction, + carte de création ;
//   · Vue Détail (selectedLibrary renseigné) : tableau des permissions de la direction et bouton de retour.
// - Pour la direction choisie : affiche les permissions (utilisateurs et groupes), permet d'en ajouter,
//   d'en supprimer, de gérer les membres des groupes SharePoint et de créer une nouvelle direction.
// - Bouton "Comment ça marche ?" (en-tête) : fenêtre d'aide pour les non-initiés (mode d'emploi,
//   différence Utilisateur / Groupe, niveaux d'accès, sécurités).
// - Garde-fou : impossible de retirer le dernier titulaire du niveau "Contrôle total".
// - Chaque action est journalisée via logAction (voir pnpjsConfig.ts).
// Prérequis : l'héritage des permissions de la bibliothèque doit avoir été rompu par l'IT (sinon erreur 400).

import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
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

// La vue Accueil affiche une carte pour chacune des bibliothèques de documents visibles dont le titre commence par ce préfixe
const DIRECTION_PREFIX = "Direction ";
const TEMPLATE_DOCUMENT_LIBRARY = 101;
const GROUPE_ASSOCIES = "Associés"; // À adapter au nom exact du groupe SharePoint
const CONTROLE_TOTAL = "Contrôle total";
const MESSAGE_GROUPE_PROTEGE = "Par sécurité, la suppression d'un groupe entier n'est pas autorisée ici.";
// Caractères refusés par SharePoint dans le nom (et l'URL) d'une bibliothèque
const CARACTERES_INTERDITS = /[~"#%&*:<>?/\\{|}]/;

// Contenu de la fenêtre d'aide : niveaux d'accès SharePoint standards, expliqués sans jargon
const NIVEAUX_AIDE: { nom: string; description: string }[] = [
  { nom: "Lecture", description: "Peut ouvrir, lire et télécharger les documents. Ne peut rien modifier ni supprimer." },
  { nom: "Afficher uniquement", description: "Peut lire les documents à l'écran, sans pouvoir les télécharger." },
  { nom: "Collaboration", description: "Peut lire, ajouter, modifier et supprimer des documents." },
  { nom: "Modification", description: "Comme « Collaboration », et peut en plus organiser la bibliothèque (dossiers, colonnes, affichages)." },
  { nom: "Conception", description: "Peut modifier les documents et l'apparence de la bibliothèque. Réservé aux besoins techniques." },
  { nom: CONTROLE_TOTAL, description: "Peut tout faire, y compris décider qui a accès. À réserver aux associés." }
];

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
  const [directionsLoaded, setDirectionsLoaded] = useState<boolean>(false);
  const [isCreateModalOpen, setIsCreateModalOpen] = useState<boolean>(false);
  const [newDirectionName, setNewDirectionName] = useState<string>("");
  const [createLoading, setCreateLoading] = useState<boolean>(false);
  const [createMessage, setCreateMessage] = useState<string>("");

  const [isHelpOpen, setIsHelpOpen] = useState<boolean>(false);

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
    } finally {
      setDirectionsLoaded(true);
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
    if (!isAddModalOpen && !isCreateModalOpen && !membersGroup && !isHelpOpen) return;
    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.key !== "Escape") return;
      // La fenêtre d'aide s'affiche par-dessus les autres modales : Échap ne ferme qu'elle
      if (isHelpOpen) {
        setIsHelpOpen(false);
        return;
      }
      if (membersGroup && !memberAddLoading && removingMemberId === null) {
        setMembersGroup(null);
        setMembers([]);
        setMemberEmailToAdd("");
        setMemberMessage("");
      }
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
  }, [isAddModalOpen, addLoading, isCreateModalOpen, createLoading, membersGroup, memberAddLoading, removingMemberId, isHelpOpen]);

  // Identifiant du dernier chargement lancé : un retour à l'accueil ou l'ouverture d'une autre direction
  // pendant un chargement rend la réponse précédente obsolète, elle est alors ignorée.
  const permissionsRequestRef = useRef<number>(0);

  const loadPermissions = async (libraryName: string): Promise<void> => {
    const requestId = ++permissionsRequestRef.current;
    setLoading(true);
    setError("");
    setSelectedLibrary(libraryName);
    try {
      const roleAssignments: any[] = await sp.web.lists
        .getByTitle(libraryName)
        .roleAssignments.expand("Member", "RoleDefinitionBindings")();
      if (requestId !== permissionsRequestRef.current) return;

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
      if (requestId !== permissionsRequestRef.current) return;
      setError("Impossible de récupérer les permissions. Vérifiez le nom exact de la bibliothèque et vos droits d'accès.");
    } finally {
      if (requestId === permissionsRequestRef.current) setLoading(false);
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

  // Pas de fermeture de la modale des membres pendant un ajout ou un retrait en cours
  const membersBusy = memberAddLoading || removingMemberId !== null;

  const closeMembersModal = (): void => {
    if (membersBusy) return;
    closeMembers();
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
    // Préfixe automatique pour que la bibliothèque apparaisse parmi les cartes de l'accueil
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

  // Vue Accueil -> Vue Détail
  const openDirection = (libraryName: string): void => {
    closeMembers();
    setSuccessMessage("");
    loadPermissions(libraryName).catch(console.error);
  };

  // Vue Détail -> Vue Accueil : on vide la sélection et on invalide un éventuel chargement en cours.
  // Bloqué pendant un retrait, car celui-ci recharge ensuite les permissions de la direction.
  const backToHome = (): void => {
    if (removingId !== null) return;
    permissionsRequestRef.current++;
    closeMembers();
    setIsAddModalOpen(false);
    setSelectedLibrary("");
    setPermissions([]);
    setLoading(false);
    setError("");
    setSuccessMessage("");
  };

  // "Direction Juridique" -> "Juridique" (libellé affiché sur les cartes)
  const libelleDirection = (dir: string): string =>
    dir.substring(DIRECTION_PREFIX.length).trim() || dir;

  const masthead = (
    <header className={styles.masthead}>
      <img className={styles.logo} src={logoCofidest} alt="Cofidest" />
      <div>
        <h1 className={styles.title}>Gestionnaire de Permissions</h1>
        <p className={styles.subtitle}>Gestion des accès aux bibliothèques par direction</p>
      </div>
      {isAuthorized && (
        <button type="button" className={styles.helpButton} onClick={() => setIsHelpOpen(true)}>
          <span className={styles.helpIcon} aria-hidden="true">?</span>
          Comment ça marche ?
        </button>
      )}
    </header>
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

  // ---------- Vue Accueil : grille de cartes ----------
  const homeView = (
    <section key="home" className={styles.view} aria-labelledby="pm-home-title">
      <div className={styles.sectionHeader}>
        <h2 id="pm-home-title" className={styles.sectionTitle}>Directions</h2>
        <p className={styles.sectionHint}>Choisissez une direction pour consulter et gérer les accès à sa bibliothèque.</p>
      </div>

      {!directionsLoaded ? (
        <p className={styles.statusText}>Chargement des directions...</p>
      ) : (
        <>
          {directions.length === 0 && (
            <p className={styles.emptyState}>{`Aucune bibliothèque « ${DIRECTION_PREFIX}… » n'a été trouvée sur ce site.`}</p>
          )}
          <div className={styles.cardGrid}>
            {directions.map((dir) => {
              const libelle = libelleDirection(dir);
              return (
                <button key={dir} type="button" className={styles.directionCard} onClick={() => openDirection(dir)}>
                  <span className={styles.cardEyebrow}>Direction</span>
                  <span className={styles.cardTitle}>{libelle}</span>
                  <span className={styles.cardAction}>
                    Gérer les accès <span className={styles.cardArrow} aria-hidden="true">→</span>
                  </span>
                </button>
              );
            })}
            <button type="button" className={`${styles.directionCard} ${styles.createCard}`} onClick={openCreateModal}>
              <span className={styles.createIcon} aria-hidden="true">+</span>
              <span className={styles.cardTitle}>Créer une nouvelle Direction</span>
              <span className={styles.cardHint}>Nouvelle bibliothèque aux permissions indépendantes</span>
            </button>
          </div>
        </>
      )}
    </section>
  );

  // ---------- Vue Détail : permissions de la direction sélectionnée ----------
  const detailView = (
    <section key="detail" className={styles.view} aria-labelledby="pm-detail-title">
      <button
        type="button"
        className={styles.backButton}
        onClick={backToHome}
        disabled={removingId !== null}
      >
        <span className={styles.backArrow} aria-hidden="true">←</span>
        Retour aux directions
      </button>

      <div className={styles.toolbar}>
        <div>
          <span className={styles.cardEyebrow}>Bibliothèque</span>
          <h2 id="pm-detail-title" className={styles.libraryTitle}>{selectedLibrary}</h2>
        </div>
        {!loading && !error && (
          <button className={`${styles.primaryButton} ${styles.addAccessButton}`} onClick={openAddModal}>
            <span className={styles.plusIcon} aria-hidden="true">+</span>
            Ajouter un accès
          </button>
        )}
      </div>

      {loading && <p className={styles.statusText}>Chargement...</p>}
      {error && <p className={styles.errorBanner}>{error}</p>}

      {!loading && !error && (
        <>
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
        </>
      )}
    </section>
  );

  return (
    <div className={styles.container}>
      {masthead}

      {selectedLibrary ? detailView : homeView}

      {membersGroup && (
        <div className={styles.modalBackdrop} onClick={closeMembersModal}>
          <div
            className={`${styles.modal} ${styles.modalWide}`}
            role="dialog"
            aria-modal="true"
            aria-labelledby="pm-members-modal-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.modalHeader}>
              <div>
                <h3 id="pm-members-modal-title" className={styles.modalTitle}>Membres du groupe « {membersGroup.nom} »</h3>
                <p className={styles.modalSubtitle}>{selectedLibrary}</p>
              </div>
              <button className={styles.closeButton} onClick={closeMembersModal} disabled={membersBusy} aria-label="Fermer">
                ×
              </button>
            </div>

            <div className={styles.modalBody}>
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

              <div className={styles.field}>
                <span className={styles.fieldLabel}>Ajouter un membre</span>
                <div className={styles.inlineForm}>
                  <input
                    type="email"
                    className={styles.input}
                    placeholder="email@cofidestsas.com"
                    value={memberEmailToAdd}
                    onChange={(e) => setMemberEmailToAdd(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter" && !memberAddLoading) handleAddMember().catch(console.error); }}
                    autoFocus
                  />
                  <button className={styles.primaryButton} onClick={handleAddMember} disabled={memberAddLoading}>
                    {memberAddLoading ? "Ajout..." : "Ajouter au groupe"}
                  </button>
                </div>
              </div>
              {memberMessage && <p className={styles.feedbackMessage}>{memberMessage}</p>}
            </div>

            <div className={styles.modalFooter}>
              <button className={styles.secondaryButton} onClick={closeMembersModal} disabled={membersBusy}>
                Fermer
              </button>
            </div>
          </div>
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

      {isHelpOpen && (
        <div className={styles.modalBackdrop} onClick={() => setIsHelpOpen(false)}>
          <div
            className={`${styles.modal} ${styles.modalWide}`}
            role="dialog"
            aria-modal="true"
            aria-labelledby="pm-help-modal-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.modalHeader}>
              <div>
                <h3 id="pm-help-modal-title" className={styles.modalTitle}>Comment ça marche ?</h3>
                <p className={styles.modalSubtitle}>Guide rapide du Gestionnaire de Permissions</p>
              </div>
              <button className={styles.closeButton} onClick={() => setIsHelpOpen(false)} aria-label="Fermer">
                ×
              </button>
            </div>

            <div className={styles.modalBody}>
              <section className={styles.helpSection}>
                <h4 className={styles.helpTitle}>À quoi sert cet outil ?</h4>
                <p className={styles.helpText}>
                  Chaque direction du cabinet possède son propre espace de documents (une « bibliothèque »).
                  Cet outil vous permet de choisir <strong>qui peut voir ou modifier</strong> les documents de chaque direction,
                  simplement, sans passer par les réglages de SharePoint.
                </p>
              </section>

              <section className={styles.helpSection}>
                <h4 className={styles.helpTitle}>Mode d’emploi en 4 étapes</h4>
                <ol className={styles.helpList}>
                  <li><strong>Cliquez sur la carte d’une direction</strong> sur la page d’accueil.</li>
                  <li>Un tableau affiche <strong>toutes les personnes et tous les groupes</strong> qui ont accès à ses documents, avec leur niveau d’accès.</li>
                  <li>Pour donner un accès, cliquez sur <strong>« + Ajouter un accès »</strong>, indiquez à qui, puis choisissez le niveau.</li>
                  <li>Pour enlever un accès, cliquez sur <strong>« Retirer »</strong> sur la ligne concernée, puis confirmez.</li>
                </ol>
                <p className={styles.helpNote}>
                  Chaque consultation, ajout ou retrait est enregistré, et les associés en sont informés par email.
                </p>
              </section>

              <section className={styles.helpSection}>
                <h4 className={styles.helpTitle}>Utilisateur ou Groupe : quelle différence ?</h4>
                <div className={styles.helpCompare}>
                  <div className={styles.helpCard}>
                    <span className={styles.typeTag}>Utilisateur</span>
                    <p className={styles.helpText}><strong>Une seule personne</strong>, reconnue par son adresse email.</p>
                    <p className={styles.helpText}>L’accès ne concerne qu’elle, et uniquement pour cette direction.</p>
                    <p className={styles.helpExample}>Exemple : donner l’accès à prenom.nom@cofidestsas.com</p>
                  </div>
                  <div className={styles.helpCard}>
                    <span className={styles.typeTag}>Groupe</span>
                    <p className={styles.helpText}><strong>Une liste de personnes</strong> réunies sous un même nom (par exemple « {GROUPE_ASSOCIES} »).</p>
                    <p className={styles.helpText}>
                      Donner un accès au groupe le donne à <strong>tous ses membres</strong>. Le bouton « Membres » permet de voir
                      qui en fait partie, d’y ajouter ou d’en retirer quelqu’un.
                    </p>
                    <p className={styles.helpExample}>Exemple : donner l’accès au groupe « Comptabilité »</p>
                  </div>
                </div>
                <p className={styles.helpWarning}>
                  <strong>Attention :</strong> un groupe est commun à toutes les directions. Ajouter une personne à un groupe
                  lui donne aussi accès à <strong>toutes les autres directions</strong> où ce groupe a déjà accès.
                  En cas de doute, ajoutez plutôt la personne en tant qu’Utilisateur.
                </p>
              </section>

              <section className={styles.helpSection}>
                <h4 className={styles.helpTitle}>Les niveaux d’accès</h4>
                <dl className={styles.helpLevels}>
                  {NIVEAUX_AIDE.map((n) => (
                    <div key={n.nom} className={styles.helpLevel}>
                      <dt className={n.nom === CONTROLE_TOTAL ? styles.levelFullControl : undefined}>{n.nom}</dt>
                      <dd>{n.description}</dd>
                    </div>
                  ))}
                </dl>
                <p className={styles.helpNote}>
                  Conseil : donnez toujours le niveau le plus faible suffisant. Pour simplement consulter des documents, choisissez « Lecture ».
                  Le niveau « Accès limité », parfois visible dans le tableau, est ajouté automatiquement par SharePoint : vous pouvez l’ignorer.
                </p>
              </section>

              <section className={styles.helpSection}>
                <h4 className={styles.helpTitle}>Les sécurités intégrées</h4>
                <ul className={styles.helpList}>
                  <li>Une confirmation vous est toujours demandée avant de retirer un accès.</li>
                  <li>Il est impossible de retirer le dernier titulaire du « {CONTROLE_TOTAL} » d’une direction.</li>
                  <li>Un groupe entier ne peut pas être retiré depuis cet outil (seuls ses membres peuvent l’être).</li>
                  <li>Vous ne pouvez pas vous retirer vous-même du groupe « {GROUPE_ASSOCIES} ».</li>
                </ul>
              </section>
            </div>

            <div className={styles.modalFooter}>
              <button className={styles.primaryButton} onClick={() => setIsHelpOpen(false)}>
                J’ai compris
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PermissionsManagerMkI;
