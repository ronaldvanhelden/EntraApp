import { Routes, Route, Navigate } from 'react-router-dom';
import { AppMsalProvider } from './auth/MsalProvider';
import { Layout } from './components/Layout';
import { Home } from './pages/Home';
import { Applications } from './pages/Applications';
import { ApplicationDetail } from './pages/ApplicationDetail';
import { EnterpriseApps } from './pages/EnterpriseApps';
import { EnterpriseAppDetail } from './pages/EnterpriseAppDetail';
import { Settings } from './pages/Settings';
import { Search } from './pages/Search';
import { Groups, Users } from './pages/Principals';
import { GroupDetail, UserDetail } from './pages/PrincipalDetail';

export default function App() {
  return (
    <AppMsalProvider>
      <Layout>
        <Routes>
          <Route path="/" element={<Home />} />
          <Route path="/applications" element={<Applications />} />
          <Route path="/applications/:id" element={<ApplicationDetail />} />
          <Route path="/enterprise-apps" element={<EnterpriseApps />} />
          <Route path="/enterprise-apps/:id" element={<EnterpriseAppDetail />} />
          <Route path="/users" element={<Users />} />
          <Route path="/users/:id" element={<UserDetail />} />
          <Route path="/groups" element={<Groups />} />
          <Route path="/groups/:id" element={<GroupDetail />} />
          <Route path="/search" element={<Search />} />
          <Route path="/settings" element={<Settings />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </Layout>
    </AppMsalProvider>
  );
}
