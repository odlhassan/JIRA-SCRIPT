from pathlib import Path
import unittest
from unittest.mock import patch

import azure_deploy_maintenance as deploy

ROOT = Path(__file__).resolve().parent.parent


class DeploymentTests(unittest.TestCase):
    def state(self, **changes):
        return dict({'deployment': 'run', 'ready': True, 'phase': 'draining', 'available': False, 'active_requests': 0, 'background_jobs': 0, 'revision': 'new'}, **changes)

    def test_drain_waits_for_both_response_and_background_job(self):
        states = [self.state(active_requests=1), self.state(background_jobs=1), self.state(), self.state()]
        with patch.object(deploy, 'get_status', side_effect=states) as status, patch.object(deploy.time, 'sleep'):
            deploy.wait_for('run')
        self.assertEqual(status.call_count, 4)

    def test_reopen_requires_target_revision_and_deployment_owner(self):
        states = [self.state(revision='old'), self.state(deployment='other'), self.state(ready=False), self.state()]
        with patch.object(deploy, 'get_status', side_effect=states) as status, patch.object(deploy.time, 'sleep'):
            deploy.wait_for('run', revision='new')
        self.assertEqual(status.call_count, 4)

    def test_timeout_fails_without_reopening(self):
        with patch.object(deploy.time, 'monotonic', side_effect=[0, 2]):
            with self.assertRaises(RuntimeError): deploy.wait_for('run', timeout=1)

    def test_bootstrap_requires_explicit_flag(self):
        with patch.object(deploy, 'get_status', side_effect=ValueError), patch.object(deploy.time, 'sleep'):
            deploy.wait_for('run', bootstrap=True)
        with patch.object(deploy, 'get_status', side_effect=ValueError), patch.object(deploy.time, 'sleep'), patch.object(deploy.time, 'monotonic', side_effect=[0, 0, 2]):
            with self.assertRaises(RuntimeError): deploy.wait_for('run', timeout=1)

    def test_workflow_gates_before_upload_and_reopens_after_revision_check(self):
        workflow = (ROOT / '.github/workflows/azure-appservice-deploy.yml').read_text()
        positions = [workflow.index(s) for s in ['name: Enable deployment maintenance', 'name: Drain app requests and report jobs', 'name: Deploy to Azure Web App', 'name: Verify new revision and reopen app', 'name: Keep app paused after deployment failure']]
        self.assertEqual(positions, sorted(positions))
        self.assertIn('staging/DEPLOYMENT_REVISION', workflow)
        self.assertIn('maintenance/game.js', workflow)


if __name__ == '__main__':
    unittest.main()
