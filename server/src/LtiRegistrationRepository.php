<?php

declare(strict_types=1);

namespace TeamsBrightspace;

use Packback\Lti1p3\Interfaces\IDatabase;
use Packback\Lti1p3\Interfaces\ILtiDeployment;
use Packback\Lti1p3\Interfaces\ILtiRegistration;
use Packback\Lti1p3\LtiDeployment;
use Packback\Lti1p3\LtiRegistration;

final class LtiRegistrationRepository implements IDatabase
{
    public function __construct(private Config $config) {}

    public function findRegistrationByIssuer(string $iss, ?string $clientId = null): ?ILtiRegistration
    {
        $settings = $this->config->ltiRegistration();
        if ($settings['issuer'] === '' || !hash_equals($settings['issuer'], $iss)
            || ($clientId !== null && !hash_equals($settings['clientId'], $clientId))) {
            return null;
        }

        $registration = new LtiRegistration($settings);
        $registration
            ->setIssuer($settings['issuer'])
            ->setClientId($settings['clientId'])
            ->setKeySetUrl($settings['keySetUrl'])
            ->setAuthLoginUrl($settings['authLoginUrl'])
            ->setKid($settings['toolKeyId'])
            ->setToolPrivateKey($this->config->ltiPrivateKey());
        if ($settings['authTokenUrl'] !== '') {
            $registration->setAuthTokenUrl($settings['authTokenUrl']);
        }
        if ($settings['authServer'] !== '') {
            $registration->setAuthServer($settings['authServer']);
        }

        return $registration;
    }

    public function findDeployment(string $iss, string $deploymentId, ?string $clientId = null): ?ILtiDeployment
    {
        $settings = $this->config->ltiRegistration();
        if ($settings['issuer'] === '' || $settings['deploymentId'] === ''
            || !hash_equals($settings['issuer'], $iss)
            || !hash_equals($settings['deploymentId'], $deploymentId)
            || ($clientId !== null && !hash_equals($settings['clientId'], $clientId))) {
            return null;
        }

        return new LtiDeployment($settings['deploymentId']);
    }
}
