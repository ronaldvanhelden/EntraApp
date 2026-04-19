import { graph } from './client';

type TokenFn = () => Promise<string>;

export interface MeResponse {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
}

export function getMe(token: TokenFn) {
  return graph<MeResponse>(token, '/me', {
    query: { $select: 'id,displayName,userPrincipalName,mail' },
  });
}
