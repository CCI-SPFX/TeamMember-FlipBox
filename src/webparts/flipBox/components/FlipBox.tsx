import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { IFlipBoxProps } from "./IFlipBoxProps";
import { spfi, SPFx } from "@pnp/sp";
import "../style.css";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface ITeamMember {
  id: number;
  name: string;
  description: string;
  linkedinUrl: string;
  xUrl: string;
  jobTitle: string;
  email: string;
  profileImage: string;
}

const FlipBox: React.FC<IFlipBoxProps> = (props) => {
  const { context, listName, linkedInIconUrl, xIconUrl } = props;

  const [members, setMembers] = useState<ITeamMember[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  const loadTeamMembers = useCallback(async (): Promise<void> => {
    if (!listName) {
      setMembers([]);
      setLoading(false);
      return;
    }

    setLoading(true);

    try {
      const sp = spfi().using(SPFx(context));

      const items: any[] = await sp.web.lists
        .getByTitle(listName)
        .items.select(
          "Id",
          "Description",
          "LinkedinURL",
          "XURL",
          "TeamMember/Title",
          "TeamMember/EMail",
          "TeamMember/JobTitle",
        )
        .expand("TeamMember")();

      console.log(items);

      const resolvedMembers: ITeamMember[] = items.map((item: any) => {
        const email = item.TeamMember?.EMail || "";

        const photo =
          email !== "" ?
            `${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(
              email,
            )}`
          : "https://i.pravatar.cc/160";

        return {
          id: item.Id,
          name: item.TeamMember?.Title || "",
          description: item.Description || "",
          linkedinUrl: item.LinkedinURL.Url || "",
          xUrl: item.XURL.Url || "",
          jobTitle: item.TeamMember?.JobTitle || "",
          email: email,
          profileImage: photo,
        };
      });

      setMembers(resolvedMembers);
      console.log("Resolved Members:", resolvedMembers);
    } catch (error) {
      console.error("Error loading team members", error);
      setMembers([]);
    } finally {
      setLoading(false);
    }
  }, [context, listName]);

  useEffect(() => {
    loadTeamMembers().catch(console.error);
  }, [loadTeamMembers]);

  if (loading) {
    return <div>Loading team members...</div>;
  }

  if (!members.length) {
    return <div>No team members found.</div>;
  }

  return (
    <div className="teamGrid">
      {members.map((member) => (
        <div className="teamCard" key={member.id}>
          <div className="cardFace cardFront">
            <img
              className="profileImage"
              src={member.profileImage}
              alt={member.name}
            />

            <h3>{member.name}</h3>

            <span>{member.jobTitle}</span>

            <a href={`mailto:${member.email}`}>{member.email}</a>
          </div>

          <div className="cardFace cardBack">
            <p>{member.description}</p>

            <div className="socialLinks">
              {member.linkedinUrl && (
                <a
                  href={member.linkedinUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  <img src={linkedInIconUrl} alt="LinkedIn" />
                </a>
              )}

              {member.xUrl && (
                <a href={member.xUrl} target="_blank" rel="noopener noreferrer">
                  <img src={xIconUrl} alt="X" />
                </a>
              )}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};

export default FlipBox;
