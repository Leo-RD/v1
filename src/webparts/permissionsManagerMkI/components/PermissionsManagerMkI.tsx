import * as React from 'react';
import { useState } from 'react';
import { IPermissionsManagerMkIProps } from './IPermissionsManagerMkIProps';  // même dossier, donc "./" et pas "../"
import { getSP } from '../pnpjsConfig';  // un seul niveau au-dessus, pas deux
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/security/list";

interface IPermissionEntry {
  principalId: number;
  nom: string;
  type: string;
  niveau: string;
}

const DIRECTIONS = ["Direction Informatique", "Direction Financiere", "Direction RH"];

const PermissionsManagerMkI: React.FC<IPermissionsManagerMkIProps> = (props) => {
  const [selectedLibrary, setSelectedLibrary] = useState<string>("");
  const [permissions, setPermissions] = useState<IPermissionEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");

  const sp = getSP(props.context);

  const loadPermissions = async (libraryName: string): Promise<void> => {
    setLoading(true);
    setError("");
    setSelectedLibrary(libraryName);
    try {
      const roleAssignments: any[] = await sp.web.lists
        .getByTitle(libraryName)
        .roleAssignments.expand("Member", "RoleDefinitionBindings")();

      const mapped: IPermissionEntry[] = roleAssignments.map((ra) => ({
        principalId: ra.PrincipalId,
        nom: ra.Member?.Title ?? "Inconnu",
        type: ra.Member?.PrincipalType === 8 ? "Groupe" : "Utilisateur",
        niveau: (ra.RoleDefinitionBindings || []).map((r: any) => r.Name).join(", ")
      }));
      setPermissions(mapped);
    } catch (e) {
      console.error(e);
      setError("Impossible de récupérer les permissions. Vérifiez le nom exact de la bibliothèque et vos droits d'accès.");
    } finally {
      setLoading(false);
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
              </tr>
            </thead>
            <tbody>
              {permissions.map((p) => (
                <tr key={p.principalId}>
                  <td>{p.nom}</td>
                  <td>{p.type}</td>
                  <td>{p.niveau}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default PermissionsManagerMkI;