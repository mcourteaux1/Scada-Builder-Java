
package elements;

// TODO: Auto-generated Javadoc
/**
 * The Class ComboItem.
 */
public class ComboItem {

	/** The key. */
	private String key;

	/** The value. */
	private String value;

	/**
	 * Instantiates a new combo item.
	 *
	 * @param key   the key
	 * @param value the value
	 */
	public ComboItem(String key, String value) {
		this.key = key;
		this.value = value;
	}

	/**
	 * To string.
	 *
	 * @return the string
	 */
	@Override
	public String toString() {
		return key;
	}

	/**
	 * Gets the key.
	 *
	 * @return the key
	 */
	public String getKey() {
		return key;
	}

	/**
	 * Gets the value.
	 *
	 * @return the value
	 */
	public String getValue() {
		return value;
	}
}