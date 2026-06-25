package org.dhatim.fastexcel;

final class LegacyProtectionHash {

    private LegacyProtectionHash() {
    }

    static String hashPassword(String password) {
        byte[] passwordCharacters = password.getBytes();
        int hash = 0;
        if (passwordCharacters.length > 0) {
            int charIndex = passwordCharacters.length;
            while (charIndex-- > 0) {
                hash = rotateHash(hash);
                hash ^= passwordCharacters[charIndex];
            }
            hash = rotateHash(hash);
            hash ^= passwordCharacters.length;
            hash ^= (0x8000 | ('N' << 8) | 'K');
        }
        return Integer.toHexString(hash & 0xffff);
    }

    private static int rotateHash(int hash) {
        return ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    }
}
